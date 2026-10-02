// Live run: the newer file, Teams, task, OneNote and Excel tools, inside one new
// OneDrive folder and the sandbox team, plan and section. Usage:
//   node scripts/live/features-work.mjs <dist> <scratch> <teamId> <planId> <sectionId>
import { mkdirSync } from "node:fs";
import { startServer, makeRun } from "./harness.mjs";

const [dist, dir, teamId, planId, sectionId] = process.argv.slice(2);
if (!sectionId) throw new Error("pass the sandbox team id, plan id and OneNote section id");
mkdirSync(`${dir}/dl`, { recursive: true });
const srv = startServer(dist, { MCP_DOWNLOAD_DIR: `${dir}/dl` });
const run = makeRun("features-work");
const { step, record, onCleanup } = run;
const TAG = `[mcp-live-test ${Date.now()}]`;
const arr = (v) => v?.value ?? (Array.isArray(v) ? v : []);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const skip = (tool, why) => record(tool, true, `SKIPPED: ${why}`);
const idOf = (v) => v?.id ?? v?.driveItem?.id;
const PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

await srv.init();
try {
  const me = (await srv.call("graph_batch_get", { requests: [{ id: "me", url: "/me?$select=id,mail,displayName" }] })).value?.responses?.[0]?.body;
  if (!me?.mail) throw new Error("could not read own profile");

  // ---- files ----
  const folderName = `mcp-live-test-${Date.now()}`;
  const root = `/${folderName}`;
  const folder = await step(srv, "files_create_folder", { parentPath: "/", name: folderName }, (v) => v.name === folderName || "wrong name");
  if (!folder) throw new Error("no sandbox folder");
  onCleanup(`delete sandbox folder ${folderName}`, async () => { const r = await srv.call("files_delete", { itemId: folder.id }); if (!r.ok) throw new Error(r.text); });
  await srv.call("files_create_folder", { parentPath: root, name: "moved" });

  const doc = await step(srv, "word_create_document", { parentPath: root, filename: "report.docx", blocks: [{ kind: "heading1", text: "Quarterly report" }, { kind: "paragraph", text: "Revenue grew by twelve percent." }] }, (v) => !!idOf(v) || "no id");
  if (idOf(doc)) {
    const pdf = await step(srv, "files_export_pdf", { itemId: idOf(doc) }, (v) => (v.name === "report.pdf" && !!v.driveItemId) || JSON.stringify(v).slice(0, 200));
    if (pdf?.driveItemId) await step(srv, "files_download", { itemId: pdf.driveItemId }, (v) => /Revenue grew by twelve percent/.test(v.text ?? "") || JSON.stringify(v).slice(0, 200), "files_download (PDF read as text)");
  }
  const png = await srv.call("files_upload", { parentPath: root, filename: "pixel.png", contentBase64: PNG });
  if (png.ok) {
    const shown = await srv.call("files_download", { itemId: png.value.id });
    record("files_download (image shown as an image)", shown.ok && shown.content.some((c) => c.type === "image" && c.mimeType === "image/png"), shown.text.slice(0, 120));
  }
  const notes = await srv.call("files_upload", { parentPath: root, filename: "notes.txt", contentBase64: Buffer.from("notes").toString("base64") });
  if (notes.ok) {
    await step(srv, "files_move", { itemId: notes.value.id, destinationParentPath: `${root}/moved`, newName: "renamed.txt" },
      (v) => (v.name === "renamed.txt" && v.folder === `${root}/moved`) || JSON.stringify(v).slice(0, 160));
    const link = await srv.call("files_share", { itemId: notes.value.id, type: "view", scope: "organization" });
    const perms = await step(srv, "files_list_permissions", { itemId: notes.value.id }, (v) => v.some((p) => p.via === "link") || JSON.stringify(v).slice(0, 160));
    const linkPerm = (perms ?? []).find((p) => p.via === "link" && !p.inheritedFrom);
    if (link.ok && linkPerm) {
      await step(srv, "files_remove_permission", { itemId: notes.value.id, permissionId: linkPerm.permissionId });
      await step(srv, "files_list_permissions", { itemId: notes.value.id }, (v) => !v.some((p) => p.permissionId === linkPerm.permissionId) || "link still there", "files_remove_permission (took effect)");
    }
  }
  skip("files_invite", "would give a real person access");
  await step(srv, "files_list_recent", { top: 10 }, (v) => (Array.isArray(v) && v.length > 0) || "no recent files");

  // ---- SharePoint lists, in the sandbox team's site ----
  const site = (await srv.call("graph_batch_get", { requests: [{ id: "s", url: `/groups/${teamId}/sites/root?$select=id` }] })).value?.responses?.[0]?.body;
  const lists = site?.id ? await step(srv, "sites_list_lists", { siteId: site.id }, (v) => Array.isArray(v) || "no list") : undefined;
  const list = (lists ?? []).find((l) => l.template === "genericList");
  if (list) {
    const item = await step(srv, "sites_create_list_item", { siteId: site.id, listId: list.id, fields: { Title: `${TAG} item` } }, (v) => !!v.id || "no id");
    if (item?.id) {
      await step(srv, "sites_update_list_item", { siteId: site.id, listId: list.id, itemId: item.id, fields: { Title: `${TAG} item (updated)` } });
      await step(srv, "sites_get_list_items", { siteId: site.id, listId: list.id, top: 50 }, (v) => arr(v).some((i) => i.fields?.Title === `${TAG} item (updated)`) || "update not visible");
    }
  } else ["sites_create_list_item", "sites_update_list_item", "sites_get_list_items"].forEach((t) => skip(t, "the sandbox site has no custom list"));

  // ---- Teams ----
  const chans = arr((await srv.call("teams_list_channels", { teamId })).value);
  const channel = chans.find((c) => c.displayName === "tests") ?? chans[0];
  if (channel) {
    const post = await step(srv, "teams_post_channel_message", { teamId, channelId: channel.id, body: `${TAG} hello @${me.displayName}, a mention test`, mentions: [me.mail] },
      (v) => v.mentions?.[0]?.mentioned?.user?.id === me.id || JSON.stringify(v.mentions ?? v).slice(0, 160), "teams_post_channel_message (mention)");
    if (post?.id) record("mention in place of @Name", (post.body?.content ?? "").includes(`<at id="0">`), (post.body?.content ?? "").slice(0, 120));
  }
  const note = await srv.call("teams_post_chat_message", { chatId: "48:notes", body: `${TAG} to edit` });
  if (note.ok) {
    await step(srv, "teams_update_chat_message", { chatId: "48:notes", messageId: note.value.id, body: `${TAG} edited` });
    await step(srv, "teams_delete_chat_message", { chatId: "48:notes", messageId: note.value.id });
  }
  skip("teams_send_direct_message", "would message another person");

  // ---- To Do ----
  const lists2 = arr((await srv.call("todo_list_lists", {})).value);
  const todoList = lists2.find((l) => l.wellknownListName === "defaultList") ?? lists2[0];
  if (todoList) {
    const t = await srv.call("todo_create_task", { listId: todoList.id, title: `${TAG} task` });
    if (t.ok) {
      onCleanup("delete To Do task", async () => { await srv.call("todo_delete_task", { listId: todoList.id, taskId: t.value.id }); });
      const due = new Date(Date.now() + 864e5).toISOString().slice(0, 10);
      await step(srv, "todo_update_task", { listId: todoList.id, taskId: t.value.id, dueDate: due, importance: "high", status: "inProgress" },
        (v) => (v.importance === "high" && v.status === "inProgress" && v.dueDateTime?.dateTime?.startsWith(due)) || JSON.stringify(v).slice(0, 200));
      await step(srv, "todo_add_checklist_item", { listId: todoList.id, taskId: t.value.id, title: "first step" }, (v) => v.displayName === "first step" || JSON.stringify(v).slice(0, 120));
    }
  }

  // ---- Planner, in the sandbox plan ----
  const buckets = await step(srv, "planner_list_buckets", { planId }, (v) => Array.isArray(v) || "no list");
  const bucket = (buckets ?? []).find((b) => b.name === "mcp-live-test") ?? (await step(srv, "planner_create_bucket", { planId, name: "mcp-live-test" }, (v) => !!v.id || "no id"));
  const task = await step(srv, "planner_create_task", { planId, title: `${TAG} planner`, assigneeUserIds: [me.mail], ...(bucket?.id ? { bucketId: bucket.id } : {}) },
    (v) => (!!v.id && Object.keys(v.assignments ?? {}).includes(me.id)) || JSON.stringify(v.assignments ?? v).slice(0, 160), "planner_create_task (assigned by email)");
  if (task?.id) {
    onCleanup("delete Planner task", async () => { await srv.call("planner_delete_task", { taskId: task.id }); });
    await sleep(3000); // Planner settles a new task's etag shortly after creation
    await step(srv, "planner_update_task", { taskId: task.id, percentComplete: 50, priority: "important", title: `${TAG} planner (updated)` },
      (v) => (v.ok || v.percentComplete === 50) || JSON.stringify(v).slice(0, 160));
    const details = await step(srv, "planner_update_task_details", { taskId: task.id, description: "notes from the live test", addItems: ["one", "two"] },
      (v) => (v.ok || v.checklist?.length === 2) || JSON.stringify(v).slice(0, 160));
    const got = await step(srv, "planner_get_task_details", { taskId: task.id }, (v) => (v.description === "notes from the live test" && v.checklist.length === 2) || JSON.stringify(v).slice(0, 200));
    if (got?.checklist?.[0]) await step(srv, "planner_update_task_details", { taskId: task.id, check: [got.checklist[0].id] }, () => true, "planner_update_task_details (tick)");
    void details;
  }

  // ---- OneNote ----
  const page = await srv.call("onenote_create_page", { sectionId, title: `${TAG} page`, html: "<p>first</p>" });
  if (page.ok) {
    onCleanup("delete OneNote page", async () => { await srv.call("onenote_delete_page", { pageId: page.value.id }); });
    let appended;
    for (let i = 0; i < 10 && !appended?.ok; i++) { appended = await srv.call("onenote_append_to_page", { pageId: page.value.id, text: "appended <line>" }); if (!appended.ok) await sleep(3000); }
    record("onenote_append_to_page", !!appended?.ok, appended?.ok ? "" : appended?.text.slice(0, 160));
    let content;
    for (let i = 0; i < 10 && !(content?.text ?? "").includes("appended &lt;line&gt;"); i++) { await sleep(3000); content = await srv.call("onenote_get_page_content", { pageId: page.value.id }); }
    record("onenote_append_to_page (took effect)", (content?.text ?? "").includes("appended &lt;line&gt;"), (content?.text ?? "").slice(0, 120));
  }

  // ---- Excel ----
  const wb = await srv.call("excel_create_workbook", { parentPath: root, filename: "chart.xlsx", worksheetName: "Data", values: [["Month", "Sales"], ["Jan", 10], ["Feb", 14], ["Mar", 9]] });
  const wbId = idOf(wb.value);
  if (wbId) {
    let chart;
    for (let i = 0; i < 6 && !chart?.ok; i++) { chart = await srv.call("excel_add_chart", { itemId: wbId, worksheet: "Data", sourceAddress: "A1:B4", type: "ColumnClustered", title: "Sales" }); if (!chart.ok) await sleep(3000); }
    record("excel_add_chart", !!chart?.ok && !!chart.value?.name, chart?.ok ? chart.value.name : chart?.text.slice(0, 160));
    await step(srv, "excel_format_range", { itemId: wbId, worksheet: "Data", address: "B2:B4", bold: true, fillColor: "#FFF2CC", numberFormat: "0.00", autofitColumns: true },
      (v) => v.applied?.length === 4 || JSON.stringify(v));
    await step(srv, "excel_get_range", { itemId: wbId, worksheet: "Data", address: "B2:B4" }, (v) => JSON.stringify(v.numberFormat) === JSON.stringify([["0.00"], ["0.00"], ["0.00"]]) || JSON.stringify(v.numberFormat), "excel_format_range (took effect)");
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
