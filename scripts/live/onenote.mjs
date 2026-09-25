// Live run: OneNote page tools and the OneNote resource, in a notebook set aside
// for testing. Usage: node scripts/live/onenote.mjs <dist> <scratch> <sectionId>
import { startServer, makeRun } from "./harness.mjs";

const [dist, dir, sectionId] = process.argv.slice(2);
if (!sectionId) throw new Error("pass the id of a section in a test notebook");
const srv = startServer(dist);
const run = makeRun("onenote");
const { step, record, onCleanup } = run;
const TAG = `[mcp-live-test ${Date.now()}]`;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

await srv.init();
try {
  await step(srv, "onenote_list_notebooks", {}, (v) => JSON.stringify(v).includes("mcp-live-test") || "test notebook not listed");
  await step(srv, "onenote_list_sections", {}, (v) => JSON.stringify(v).includes(sectionId) || "test section not listed");
  const page = await step(srv, "onenote_create_page", {
    sectionId, title: `${TAG} page`,
    html: "<h1>Heading</h1><p>First&nbsp;paragraph</p><ul><li>item one</li><li>item two</li></ul>",
  }, (v) => !!v.id || "no id");
  if (!page?.id) throw new Error("page not created");
  let deleted = false;
  onCleanup("delete test page", async () => { if (!deleted) await srv.call("onenote_delete_page", { pageId: page.id }); });

  // OneNote indexes a new page asynchronously; content and listings can lag.
  let content;
  for (let i = 0; i < 10 && !content?.ok; i++) { content = await srv.call("onenote_get_page_content", { pageId: page.id }); if (!content.ok) await sleep(3000); }
  record("onenote_get_page_content", !!content?.ok && content.text.includes("item two"), content?.ok ? "" : content?.text.slice(0, 200));
  let listed;
  for (let i = 0; i < 10 && !listed; i++) { const r = await srv.call("onenote_list_pages", { sectionId }); listed = r.ok && r.text.includes(page.id); if (!listed) await sleep(3000); }
  record("onenote_list_pages", !!listed, listed ? "" : "new page never listed");

  // The resource: plain text, list structure kept, markup gone.
  const res = await srv.readResource(`m365://onenote/${encodeURIComponent(page.id)}`);
  const text = res?.contents?.[0]?.text ?? "";
  record("resource m365://onenote/{id}", text.includes("Heading") && text.includes("- item one\n- item two") && !text.includes("<"), JSON.stringify(text).slice(0, 200));

  deleted = !!(await step(srv, "onenote_delete_page", { pageId: page.id }));
} catch (e) {
  console.log("ABORTED:", e.message);
} finally {
  await run.cleanup();
  run.save(dir);
  const failed = run.results.filter((r) => !r.pass);
  console.log(`\n${run.results.length - failed.length}/${run.results.length} passed`);
  srv.kill();
}
