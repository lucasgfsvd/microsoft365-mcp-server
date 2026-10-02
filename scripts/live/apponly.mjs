// Live run: app-only (client-credentials) mode, working for MCP_USER.
// Needs an app registration with application permissions and admin consent; the
// settings come from the environment, so no secret is ever on the command line:
//   MCP_TENANT_ID, MCP_CLIENT_ID, MCP_USER, and MCP_CLIENT_CERTIFICATE_PATH (or MCP_CLIENT_SECRET)
// Usage: node scripts/live/apponly.mjs <dist> <scratch>
// Everything it creates is tagged and deleted; it sends no mail.
import { startServer, makeRun } from "./harness.mjs";

const [dist, dir] = process.argv.slice(2);
const user = process.env.MCP_USER;
if (!user || !process.env.MCP_TENANT_ID || !process.env.MCP_CLIENT_ID) throw new Error("set MCP_TENANT_ID, MCP_CLIENT_ID and MCP_USER");
const srv = startServer(dist, { MCP_AUTH_MODE: "client-credentials" });
const run = makeRun("apponly");
const { step, record, onCleanup } = run;
const TAG = `[mcp-live-test ${Date.now()}]`;
const arr = (v) => v?.value ?? (Array.isArray(v) ? v : []);

await srv.init();
try {
  const names = new Set((await srv.listTools()).map((t) => t.name));
  const hidden = ["onenote_list_notebooks", "teams_post_chat_message", "graph_search"].filter((n) => names.has(n));
  record("delegated-only tools hidden", hidden.length === 0 && names.has("mail_list_messages"), hidden.join(", "));
  await step(srv, "auth_status", {}, (v) => v.signedIn === true || `not signed in: ${JSON.stringify(v)}`);

  // /me is MCP_USER, in direct calls and inside $batch alike.
  await step(srv, "graph_batch_get", { requests: [{ id: "me", url: "/me?$select=userPrincipalName,mail" }] },
    (v) => [v.responses?.[0]?.body?.userPrincipalName, v.responses?.[0]?.body?.mail].some((x) => x?.toLowerCase() === user.toLowerCase()) || JSON.stringify(v.responses?.[0]).slice(0, 200),
    "graph_batch_get (/me is MCP_USER)");

  // ---- mail ----
  const inbox = await step(srv, "mail_list_messages", { top: 3 }, (v) => Array.isArray(arr(v)) || "no list");
  const first = arr(inbox)[0];
  if (first) await step(srv, "mail_get_message", { id: first.id }, (v) => v.id === first.id || "id mismatch");
  const draft = await step(srv, "mail_create_draft", { to: [{ address: user }], subject: `${TAG} draft`, body: "app-only draft" }, (v) => !!v.id || "no id");
  if (draft?.id) await step(srv, "mail_delete_message", { id: draft.id });
  const refused = await srv.call("mail_list_messages", { mailbox: "someone.else@example.com" });
  record("mailbox argument refused app-only", !refused.ok && /MCP_USER only/.test(refused.text), refused.text.slice(0, 160));

  // ---- calendar ----
  await step(srv, "calendar_list_calendars", {}, (v) => arr(v).length > 0 || "no calendars");
  const day = new Date(Date.now() + 30 * 864e5).toISOString().slice(0, 10);
  const ev = await step(srv, "calendar_create_event", {
    subject: `${TAG} event`, start: { dateTime: `${day}T09:00:00`, timeZone: "UTC" }, end: { dateTime: `${day}T09:30:00`, timeZone: "UTC" },
  }, (v) => !!v.id || "no id");
  if (ev?.id) await step(srv, "calendar_delete_event", { id: ev.id });
  await step(srv, "calendar_get_free_busy", { schedules: [user], startDateTime: `${day}T08:00:00`, endDateTime: `${day}T18:00:00`, timeZone: "UTC" }, (v) => JSON.stringify(v).includes("availabilityView") || "no availability");

  // ---- files, Excel ----
  const folderName = `mcp-live-test-${Date.now()}`;
  const folder = await step(srv, "files_create_folder", { parentPath: "/", name: folderName }, (v) => v.name === folderName || "wrong name");
  if (folder?.id) {
    onCleanup("delete sandbox folder", async () => { await srv.call("files_delete", { itemId: folder.id }); });
    await step(srv, "files_upload", { parentPath: `/${folderName}`, filename: "a.txt", contentBase64: Buffer.from("app-only").toString("base64") }, (v) => v.size > 0 || "no size");
    const wb = await step(srv, "excel_create_workbook", { parentPath: `/${folderName}`, filename: "a.xlsx", worksheetName: "Data", values: [["x"], [1]] }, (v) => !!(v.id ?? v.driveItem?.id) || "no id");
    const wbId = wb?.id ?? wb?.driveItem?.id;
    if (wbId) await step(srv, "excel_get_range", { itemId: wbId, worksheet: "Data", address: "A1:A2" }, (v) => JSON.stringify(v).includes("x") || "no values");
  }
  const resources = await srv.listResources();
  record("resources/list (recent items for MCP_USER)", resources.length > 0, `${resources.length} entries`);

  // ---- tasks, people, Teams ----
  const lists = await step(srv, "todo_list_lists", {}, (v) => arr(v).length > 0 || "no lists");
  const list = arr(lists)[0];
  if (list) {
    const t = await step(srv, "todo_create_task", { listId: list.id, title: `${TAG} task` }, (v) => !!v.id || "no id");
    if (t?.id) await step(srv, "todo_delete_task", { listId: list.id, taskId: t.id });
  }
  await step(srv, "contacts_list", { top: 3 }, (v) => Array.isArray(arr(v)) || "no list");
  await step(srv, "contacts_people_search", { query: "a", top: 3 }, (v) => Array.isArray(arr(v)) || "no list");
  await step(srv, "teams_list_joined", {}, (v) => Array.isArray(arr(v)) || "no list");
  await step(srv, "planner_list_plans", {}, (v) => Array.isArray(arr(v)) || "no list");
} catch (e) {
  console.log("ABORTED:", e.message);
} finally {
  await run.cleanup();
  run.save(dir);
  const failed = run.results.filter((r) => !r.pass);
  console.log(`\n${run.results.length - failed.length}/${run.results.length} passed`);
  srv.kill();
}
