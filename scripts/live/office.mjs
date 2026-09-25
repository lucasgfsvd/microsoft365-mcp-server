// Live run: files, Word, PowerPoint, Excel — all inside one throwaway OneDrive folder.
import { execFileSync } from "node:child_process";
import { mkdirSync } from "node:fs";
import { startServer, makeRun } from "./harness.mjs";

const [dist, dir] = process.argv.slice(2);
const dl = `${dir}/dl`;
mkdirSync(dl, { recursive: true });
const srv = startServer(dist, { MCP_DOWNLOAD_DIR: dl });
const run = makeRun("office");
const { step, record, onCleanup } = run;
const idOf = (v) => v?.id ?? v?.driveItem?.id;
const has = (v, s) => JSON.stringify(v).includes(s) || `missing "${s}" in ${JSON.stringify(v).slice(0, 200)}`;
const pyCheck = (file, code) => {
  try { return execFileSync("python", ["-c", code, file], { encoding: "utf8" }).trim(); } catch (e) { return `PARSE FAILED: ${(e.stderr || e.message).toString().split("\n").slice(-3).join(" ")}`; }
};
const saveAndParse = async (tool, itemId, code) => {
  const d = await srv.call("files_download", { itemId, saveToDisk: true });
  if (!d.ok) return record(tool, false, "download for parse failed: " + d.text.slice(0, 150));
  const out = pyCheck(d.value.path, code);
  record(tool, !out.startsWith("PARSE FAILED"), out.slice(0, 250));
};
const paragraphs = async (itemId) => {
  const r = await srv.call("word_list_paragraphs", { itemId });
  const arr = Array.isArray(r.value) ? r.value : (r.value?.paragraphs ?? []);
  return arr.map((p) => (typeof p === "string" ? p : p.text));
};

await srv.init();
const folderName = `mcp-live-test-${Date.now()}`;
const root = `/${folderName}`;
try {
  // ---------------- files ----------------
  const folder = await step(srv, "files_create_folder", { parentPath: "/", name: folderName }, (v) => v.name === folderName || "wrong name");
  if (!folder) throw new Error("no sandbox folder");
  onCleanup(`delete sandbox folder ${folderName}`, async () => { const r = await srv.call("files_delete", { itemId: folder.id }); if (!r.ok) throw new Error(r.text); });

  await step(srv, "files_list_drives", {}, (v) => (v.value?.length ?? 0) > 0 || "no drives");
  await step(srv, "files_get_item", { path: root }, (v) => v.id === folder.id || "id mismatch");
  const up = await step(srv, "files_upload", { parentPath: root, filename: "notes.txt", contentBase64: Buffer.from("hello live test ✓").toString("base64") }, (v) => v.size > 0 || "no size");
  await step(srv, "files_list_children", { path: root }, (v) => has(v, "notes.txt"));
  await step(srv, "files_copy", { itemId: up?.id, destinationParentPath: root, destinationName: "notes-copy.txt" }, (v) => has(v, "notes-copy.txt"));
  await step(srv, "files_download", { itemId: up?.id }, (v) => (v.encoding === "utf8" && v.text === "hello live test ✓") || `got ${JSON.stringify(v).slice(0, 120)}`);
  await step(srv, "files_share", { itemId: up?.id, type: "view", scope: "organization" }, (v) => has(v, "webUrl"));
  await step(srv, "files_search", { query: "notes" }, (v) => Array.isArray(v.value) || "no value array");
  await step(srv, "sites_search", { query: "*" }, (v) => Array.isArray(v.value) || "no value array");

  // ---------------- Word ----------------
  const doc = await step(srv, "word_create_document", {
    parentPath: root, filename: "live.docx",
    blocks: [
      { kind: "title", text: "Live Test Title" },
      { kind: "heading1", text: "First Section" },
      { kind: "paragraph", text: "Alpha PLACEHOLDER paragraph." },
      { kind: "bullet", text: "Bullet one" },
      { kind: "bullet", text: "Bullet two" },
    ],
  }, (v) => !!idOf(v) || `no id in ${JSON.stringify(v).slice(0, 120)}`);
  if (idOf(doc)) {
    const id = idOf(doc);
    await step(srv, "word_read_text", { itemId: id }, (v) => has(v, "Alpha PLACEHOLDER"));
    let ps = await paragraphs(id);
    record("word_list_paragraphs", ps.length >= 5, `${ps.length} paragraphs: ${JSON.stringify(ps).slice(0, 160)}`);
    await step(srv, "word_replace_text", { itemId: id, replacements: [{ find: "PLACEHOLDER", replace: "Replaced" }] });
    await step(srv, "word_append_paragraph", { itemId: id, text: "Appended paragraph." });
    await step(srv, "word_append_heading", { itemId: id, text: "Appended Heading", level: 2 });
    await step(srv, "word_append_bullets", { itemId: id, items: ["Appended bullet A", "Appended bullet B"] });
    await step(srv, "word_insert_paragraph_at", { itemId: id, text: "Inserted after title.", after: 0 }); // 0 = at the top
    ps = await paragraphs(id);
    const want = ["Inserted after title.", "Live Test Title", "First Section", "Alpha Replaced paragraph.", "Bullet one", "Bullet two", "Appended paragraph.", "Appended Heading", "Appended bullet A", "Appended bullet B"];
    const same = JSON.stringify(ps) === JSON.stringify(want);
    record("word edits (order check)", same, same ? "" : `got ${JSON.stringify(ps)}`);
    const idx = ps.indexOf("Inserted after title.");
    await step(srv, "word_delete_paragraph", { itemId: id, paragraphIndex: idx + 1 }); // 1-based
    ps = await paragraphs(id);
    record("word_delete_paragraph (check)", !ps.includes("Inserted after title.") && ps.length === want.length - 1, JSON.stringify(ps).slice(0, 200));
    await saveAndParse("word (python-docx parse)", id,
      "import sys, docx; d = docx.Document(sys.argv[1]); print(len(d.paragraphs), 'paras; styles:', sorted({p.style.name for p in d.paragraphs if p.style is not None}))");
    await step(srv, "word_create_from_template", {
      templateItemId: id, parentPath: root, filename: "from-template.docx",
      replacements: [{ find: "Replaced", replace: "Templated" }],
      appendBlocks: [{ kind: "paragraph", text: "Added from template call." }],
    }, (v) => !!idOf(v) || `no id in ${JSON.stringify(v).slice(0, 120)}`);
    const tpl = await srv.call("files_get_item", { path: `${root}/from-template.docx` });
    if (tpl.ok) {
      const t = await srv.call("word_read_text", { itemId: tpl.value.id });
      record("word_create_from_template (check)", t.ok && t.text.includes("Alpha Templated") && t.text.includes("Added from template call."), t.text.slice(0, 160));
    }
  }

  // ---------------- PowerPoint ----------------
  const deck = await step(srv, "powerpoint_create_deck", {
    parentPath: root, filename: "live.pptx", title: "Live deck", author: "live test",
    slides: [
      { title: "Slide One", bullets: ["Point A", "TOKEN here"], notes: "Notes one" },
      { title: "Slide Two", bullets: ["Point B"] },
    ],
  }, (v) => !!idOf(v) || `no id in ${JSON.stringify(v).slice(0, 120)}`);
  if (idOf(deck)) {
    const id = idOf(deck);
    const count = async () => { const r = await srv.call("powerpoint_list_slides", { itemId: id }); const a = Array.isArray(r.value) ? r.value : (r.value?.slides ?? []); return a; };
    let slides = await count();
    record("powerpoint_list_slides", slides.length === 2, JSON.stringify(slides).slice(0, 200));
    await step(srv, "powerpoint_get_slide_text", { itemId: id, slideIndex: 1 }, (v) => has(v, "Slide One"));
    await step(srv, "powerpoint_extract_all_text", { itemId: id }, (v) => has(v, "Slide Two"));
    await step(srv, "powerpoint_replace_text", { itemId: id, replacements: [{ find: "TOKEN", replace: "REPLACED" }] });
    await step(srv, "powerpoint_get_slide_text", { itemId: id, slideIndex: 1 }, (v) => has(v, "REPLACED here"));
    await step(srv, "powerpoint_add_slide", { itemId: id, title: "Slide Three", bullets: ["Point C"], notes: "Notes three" });
    slides = await count();
    record("powerpoint_add_slide (check)", slides.length === 3, JSON.stringify(slides).slice(0, 200));
    await step(srv, "powerpoint_delete_slide", { itemId: id, slideIndex: 2 }); // removes "Slide Two"
    const all = await srv.call("powerpoint_extract_all_text", { itemId: id });
    record("powerpoint_delete_slide (check)", all.ok && !all.text.includes("Slide Two") && all.text.includes("Slide Three") && (await count()).length === 2, all.text.slice(0, 200));
    await saveAndParse("powerpoint (python-pptx parse)", id,
      "import sys, pptx; p = pptx.Presentation(sys.argv[1]); t = [s.shapes.title.text if s.shapes.title else None for s in p.slides]; print(len(p.slides), 'slides:', t); sys.exit(1 if None in t else 0)");
    await step(srv, "powerpoint_create_from_template", {
      templateItemId: id, parentPath: root, filename: "from-template.pptx",
      replacements: [{ find: "Point A", replace: "Point Templated" }],
      appendSlides: [{ title: "Appended From Template", bullets: ["x"] }],
    }, (v) => !!idOf(v) || `no id in ${JSON.stringify(v).slice(0, 120)}`);
    const tpl = await srv.call("files_get_item", { path: `${root}/from-template.pptx` });
    if (tpl.ok) {
      const t = await srv.call("powerpoint_extract_all_text", { itemId: tpl.value.id });
      record("powerpoint_create_from_template (check)", t.ok && t.text.includes("Point Templated") && t.text.includes("Appended From Template"), t.text.slice(0, 160));
    }
  }

  // ---------------- Excel ----------------
  const wb = await step(srv, "excel_create_workbook", {
    parentPath: root, filename: "live.xlsx", worksheetName: "Data",
    values: [["Name", "Qty"], ["apples", 3], ["pears", 4]],
  }, (v) => !!idOf(v) || `no id in ${JSON.stringify(v).slice(0, 120)}`);
  if (idOf(wb)) {
    const id = idOf(wb);
    // The Excel service can lag a fresh upload; give it a few tries.
    let ws;
    for (let i = 0; i < 6 && !(ws?.ok); i++) { ws = await srv.call("excel_list_worksheets", { itemId: id }); if (!ws.ok) await new Promise((r) => setTimeout(r, 3000)); }
    record("excel_list_worksheets", !!ws?.ok && ws.text.includes("Data"), ws?.ok ? "" : ws?.text.slice(0, 200));
    const sess = await step(srv, "excel_create_session", { itemId: id, persistChanges: true }, (v) => !!v.id || has(v, "id"));
    const sessionId = sess?.id;
    if (sessionId) onCleanup("close excel session", async () => { await srv.call("excel_close_session", { itemId: id, sessionId }); });
    const s = sessionId ? { sessionId } : {};
    await step(srv, "excel_get_range", { itemId: id, ...s, worksheet: "Data", address: "A1:B3" }, (v) => has(v, "apples"));
    await step(srv, "excel_update_range", { itemId: id, ...s, worksheet: "Data", address: "A4:B4", values: [["plums", 5]] });
    await step(srv, "excel_set_formula", { itemId: id, ...s, worksheet: "Data", address: "B5", formulas: [["=SUM(B2:B4)"]] });
    await step(srv, "excel_run_workbook_calculation", { itemId: id, ...s, calculationType: "Recalculate" });
    await step(srv, "excel_get_range", { itemId: id, ...s, worksheet: "Data", address: "B5" }, (v) => JSON.stringify(v.values ?? v).includes("12") || `B5 = ${JSON.stringify(v.values ?? v).slice(0, 80)}`);
    await step(srv, "excel_add_worksheet", { itemId: id, ...s, name: "Scratch" }, (v) => has(v, "Scratch"));
    await step(srv, "excel_rename_worksheet", { itemId: id, ...s, worksheet: "Scratch", newName: "Renamed" }, (v) => has(v, "Renamed"));
    await step(srv, "excel_delete_worksheet", { itemId: id, ...s, worksheet: "Renamed" });
    await step(srv, "excel_create_table", { itemId: id, ...s, address: "Data!A1:B4", hasHeaders: true, tableName: "Fruit" }, (v) => has(v, "Fruit"));
    await step(srv, "excel_list_tables", { itemId: id, ...s }, (v) => has(v, "Fruit"));
    await step(srv, "excel_add_table_rows", { itemId: id, ...s, tableName: "Fruit", values: [["figs", 6]] });
    await step(srv, "excel_get_table_rows", { itemId: id, ...s, tableName: "Fruit" }, (v) => has(v, "figs"));
    await step(srv, "excel_clear_range", { itemId: id, ...s, worksheet: "Data", address: "D1:D5", applyTo: "Contents" });
    if (sessionId) await step(srv, "excel_close_session", { itemId: id, sessionId });
    await saveAndParse("excel (openpyxl parse)", id,
      "import sys, openpyxl; wb = openpyxl.load_workbook(sys.argv[1]); ws = wb['Data']; print(wb.sheetnames, [[c.value for c in r] for r in ws.iter_rows(max_row=6)])");
    await step(srv, "excel_create_from_template", { templateItemId: id, parentPath: root, filename: "from-template.xlsx" }, (v) => !!idOf(v) || `no id in ${JSON.stringify(v).slice(0, 120)}`);
  }
} catch (e) {
  console.log("ABORTED:", e.message);
} finally {
  await run.cleanup();
  run.save(dir);
  const failed = run.results.filter((r) => !r.pass);
  console.log(`\n${run.results.length - failed.length}/${run.results.length} passed`);
  srv.kill();
}
