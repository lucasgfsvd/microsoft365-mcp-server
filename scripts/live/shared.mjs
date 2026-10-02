// Live run: mail and calendar tools against another mailbox (a shared mailbox,
// or a calendar the user is a delegate of), with MCP_ENABLE_SHARED_MAILBOXES.
// Needs a person to enter one device code, consenting to the .Shared scopes,
// unless MCP_TOKEN_CACHE_PATH names a cache already signed in with them (login.mjs).
// Usage: node scripts/live/shared.mjs <dist> <scratch> <mailbox-address>
// Otherwise signs in to a throwaway encrypted cache under <scratch>, removed at the end.
// Creates one draft in that mailbox and deletes it; sends nothing.
import { randomBytes } from "node:crypto";
import { mkdirSync, rmSync } from "node:fs";
import path from "node:path";
import { startServer, makeRun } from "./harness.mjs";

const [dist, scratch, mailbox] = process.argv.slice(2);
if (!mailbox) throw new Error("pass the address of a mailbox you can open");
const own = !process.env.MCP_TOKEN_CACHE_PATH;
const dir = path.resolve(scratch, `shared-${Date.now()}`);
if (own) mkdirSync(dir, { recursive: true });
const srv = startServer(dist, {
  MCP_ENABLE_SHARED_MAILBOXES: "1",
  ...(own ? { MCP_TOKEN_CACHE_PATH: path.join(dir, "tokencache.json"), MCP_TOKEN_CACHE_KEY: randomBytes(32).toString("base64") } : {}),
});
const run = makeRun("shared");
const { step, record, onCleanup } = run;
const TAG = `[mcp-live-test ${Date.now()}]`;
const arr = (v) => v?.value ?? (Array.isArray(v) ? v : []);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

await srv.init();
try {
  const tools = await srv.listTools();
  const schemaOf = (n) => tools.find((t) => t.name === n)?.inputSchema?.properties ?? {};
  record("mailbox offered on mail and calendar only", "mailbox" in schemaOf("mail_list_messages") && "mailbox" in schemaOf("calendar_list_events") && !("mailbox" in schemaOf("files_upload")));

  let done = (await srv.call("auth_status")).value?.signedIn === true;
  if (!done) {
    const s = await srv.call("auth_sign_in");
    if (!s.ok || !s.value?.userCode) throw new Error(`no device code: ${s.text.slice(0, 200)}`);
    console.log(`\n>>> SIGN IN: open ${s.value.verificationUri} and enter ${s.value.userCode}\n`);
    for (let i = 0; i < 168 && !done; i++) { await sleep(5000); done = (await srv.call("auth_status")).value?.signedIn === true; }
  }
  record("signed in with the .Shared scopes", done);
  if (!done) throw new Error("no sign-in");

  await step(srv, "mail_list_folders", { mailbox }, (v) => JSON.stringify(v).includes("Inbox") || "no Inbox", "mail_list_folders (mailbox)");
  await step(srv, "mail_list_messages", { mailbox, top: 5 }, (v) => Array.isArray(arr(v)) || "no list", "mail_list_messages (mailbox)");
  const now = new Date();
  await step(srv, "calendar_list_events", { mailbox, startDateTime: now.toISOString(), endDateTime: new Date(+now + 7 * 864e5).toISOString() },
    (v) => Array.isArray(arr(v)) || "no list", "calendar_list_events (mailbox)");

  const draft = await step(srv, "mail_create_draft", { mailbox, to: [{ address: mailbox }], subject: `${TAG} draft`, body: "in the other mailbox" }, (v) => !!v.id || "no id", "mail_create_draft (mailbox)");
  if (draft?.id) {
    onCleanup("delete draft in the other mailbox", async () => { await srv.call("mail_delete_message", { mailbox, id: draft.id }); });
    await step(srv, "mail_get_message", { mailbox, id: draft.id }, (v) => v.subject === `${TAG} draft` || `subject ${v.subject}`, "mail_get_message (mailbox)");
  }

  const refused = await srv.call("files_list_drives", { mailbox });
  record("mailbox refused on a files tool", !refused.ok && /takes no mailbox/.test(refused.text), refused.text.slice(0, 160));
  await step(srv, "mail_list_messages", { top: 1 }, (v) => Array.isArray(arr(v)) || "no list", "mail_list_messages (own mailbox still)");
} catch (e) {
  console.log("ABORTED:", e.message);
} finally {
  await run.cleanup();
  run.save(scratch);
  srv.kill();
  await sleep(500);
  if (own) rmSync(dir, { recursive: true, force: true });
  const failed = run.results.filter((r) => !r.pass);
  console.log(`\n${run.results.length - failed.length}/${run.results.length} passed`);
}
