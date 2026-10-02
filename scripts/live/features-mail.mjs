// Live run: the newer mail, calendar and people tools. Mail goes only to the
// signed-in user; events have no attendees; out-of-office is set for a date
// years ahead and put back as it was. Usage:
//   node scripts/live/features-mail.mjs <dist> <scratch>
import { createHash, randomBytes } from "node:crypto";
import { mkdirSync, writeFileSync } from "node:fs";
import { startServer, makeRun } from "./harness.mjs";

const [dist, dir] = process.argv.slice(2);
const upDir = `${dir}/up`;
const dlDir = `${dir}/dl`;
mkdirSync(upDir, { recursive: true });
mkdirSync(dlDir, { recursive: true });
const srv = startServer(dist, { MCP_UPLOAD_DIR: upDir, MCP_DOWNLOAD_DIR: dlDir });
const run = makeRun("features-mail");
const { step, record, onCleanup } = run;
const TAG = `[mcp-live-test ${Date.now()}]`;
const arr = (v) => v?.value ?? (Array.isArray(v) ? v : []);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const skip = (tool, why) => record(tool, true, `SKIPPED: ${why}`);
const b64 = (s) => Buffer.from(s).toString("base64");
const startedAt = new Date(Date.now() - 60_000).toISOString();

/** Wait for a message whose subject contains `text` to reach the inbox. */
async function arrival(text) {
  for (let i = 0; i < 25; i++) {
    await sleep(4000);
    const r = await srv.call("mail_list_messages", { receivedAfter: startedAt, top: 25 });
    const m = arr(r.value).find((x) => x.subject?.includes(text));
    if (m) return m;
  }
  return undefined;
}

await srv.init();
try {
  const me = await srv.call("graph_batch_get", { requests: [{ id: "me", url: "/me?$select=mail,userPrincipalName,displayName" }] });
  const self = me.value?.responses?.[0]?.body?.mail ?? me.value?.responses?.[0]?.body?.userPrincipalName;
  if (!self) throw new Error("could not read own address");
  onCleanup("delete tagged mail", async () => {
    for (let pass = 0; pass < 3; pass++) {
      const r = await srv.call("mail_search_messages", { query: `"${TAG}"`, top: 25 });
      for (const m of arr(r.value)) if (m.subject?.includes(TAG)) await srv.call("mail_delete_message", { id: m.id });
      await sleep(3000);
    }
  });

  // ---- attachments: one inline, one from disk large enough for an upload session ----
  const big = randomBytes(Math.round(4.5 * 1024 * 1024));
  writeFileSync(`${upDir}/big-attach.bin`, big);
  await step(srv, "mail_send_message", {
    to: [{ address: self }], subject: `${TAG} attachments`, body: "two attachments",
    attachments: [{ name: "notes.txt", contentBase64: b64("hello attachment") }, { localPath: "big-attach.bin" }],
  }, (v) => (v.ok && v.attached?.length === 2) || JSON.stringify(v).slice(0, 200), "mail_send_message (attachments)");
  const withFiles = await arrival(`${TAG} attachments`);
  record("mail with attachments arrived", !!withFiles);
  if (withFiles) {
    const atts = arr((await srv.call("mail_list_attachments", { messageId: withFiles.id })).value);
    const txt = atts.find((a) => a.name === "notes.txt");
    const bin = atts.find((a) => a.name === "big-attach.bin");
    await step(srv, "mail_get_attachment", { messageId: withFiles.id, attachmentId: txt?.id ?? "none" }, (v) => v.text === "hello attachment" || JSON.stringify(v).slice(0, 200));
    await step(srv, "mail_get_attachment", { messageId: withFiles.id, attachmentId: bin?.id ?? "none", saveToDisk: true },
      (v) => v.sha256 === createHash("sha256").update(big).digest("hex") || `sha ${v.sha256}`, "mail_get_attachment (4.5 MB to disk)");

    // ---- mark, flag, categorise, file away and back ----
    await step(srv, "mail_update_message", { id: withFiles.id, isRead: false, flag: "flagged", categories: ["mcp-live-test"] });
    await step(srv, "mail_get_message", { id: withFiles.id }, (v) => (v.isRead === false && v.flag?.flagStatus === "flagged" && v.categories?.includes("mcp-live-test")) || JSON.stringify({ r: v.isRead, f: v.flag, c: v.categories }),
      "mail_update_message (took effect)");
    const archived = await step(srv, "mail_move_message", { id: withFiles.id, destination: "archive" }, (v) => (!!v.id && v.id !== withFiles.id) || "no new id");
    if (archived?.id) await step(srv, "mail_move_message", { id: archived.id, destination: "inbox" }, (v) => !!v.id || "no id", "mail_move_message (back to inbox)");
  }

  // ---- forward as a draft, then send the draft; reply with line breaks ----
  const plain = `${TAG} original`;
  await srv.call("mail_send_message", { to: [{ address: self }], subject: plain, body: "the original message" });
  const original = await arrival(plain);
  if (original) {
    const fwd = await step(srv, "mail_forward_message", { id: original.id, to: [{ address: self }], comment: "fwd line one\nfwd line two", draft: true },
      (v) => (v.isDraft === true && /^FW/i.test(v.subject)) || JSON.stringify(v).slice(0, 200), "mail_forward_message (draft)");
    if (fwd?.id) {
      const body = (await srv.call("mail_get_message", { id: fwd.id })).value?.body?.content ?? "";
      record("forward text kept as written", /fwd line one\r?\n\s*fwd line two/.test(body) && body.includes("the original message"), JSON.stringify(body.slice(0, 160)));
      await step(srv, "mail_send_draft", { id: fwd.id });
      record("forwarded draft arrived", !!(await arrival(`FW: ${plain}`)));
    }
    await step(srv, "mail_reply_message", { id: original.id, comment: "reply line one\nreply line two" });
    const reply = await arrival(`RE: ${plain}`);
    const replyBody = reply ? (await srv.call("mail_get_message", { id: reply.id })).value?.body?.content ?? "" : "";
    record("mail_reply_message (line breaks kept)", /reply line one\r?\n\s*reply line two/.test(replyBody), JSON.stringify(replyBody.slice(0, 160)));
  } else record("original arrived", false);

  // ---- out-of-office: scheduled years ahead, then put back exactly ----
  const before = await step(srv, "mail_get_automatic_replies", {}, (v) => !!v.status || "no status");
  if (before?.status) {
    const restore = { status: before.status, ...(before.status === "scheduled" ? { start: before.scheduledStartDateTime, end: before.scheduledEndDateTime } : {}) };
    const emptyMessage = !before.internalReplyMessage;
    onCleanup("restore out-of-office", async () => {
      const r = await srv.call("mail_set_automatic_replies", { ...restore, ...(emptyMessage ? { internalMessage: "" } : {}) });
      if (!r.ok) throw new Error(r.text);
    });
    await step(srv, "mail_set_automatic_replies", {
      status: "scheduled", start: { dateTime: "2030-01-01T09:00:00", timeZone: "UTC" }, end: { dateTime: "2030-01-02T09:00:00", timeZone: "UTC" },
      ...(emptyMessage ? { internalMessage: "Away <testing>\nback soon" } : {}),
    }, (v) => (v.status === "scheduled" && (!emptyMessage || /Away &lt;testing&gt;<br>back soon/.test(v.internalReplyMessage ?? ""))) || JSON.stringify(v).slice(0, 200));
  }

  // ---- calendar ----
  const day = new Date(Date.now() + 40 * 864e5).toISOString().slice(0, 10);
  const series = await step(srv, "calendar_create_event", {
    subject: `${TAG} weekly`, start: { dateTime: `${day}T08:00:00`, timeZone: "UTC" }, end: { dateTime: `${day}T08:30:00`, timeZone: "UTC" },
    recurrence: { pattern: "weekly", occurrences: 3 },
  }, (v) => (v.recurrence?.pattern?.type === "weekly" && v.recurrence?.range?.numberOfOccurrences === 3) || JSON.stringify(v.recurrence), "calendar_create_event (recurrence)");
  if (series?.id) await step(srv, "calendar_delete_event", { id: series.id });
  const toCancel = await srv.call("calendar_create_event", { subject: `${TAG} to cancel`, start: { dateTime: `${day}T10:00:00`, timeZone: "UTC" }, end: { dateTime: `${day}T10:30:00`, timeZone: "UTC" } });
  if (toCancel.ok) {
    onCleanup("delete event if cancel left it", async () => { await srv.call("calendar_delete_event", { id: toCancel.value.id }); });
    await step(srv, "calendar_cancel_event", { id: toCancel.value.id, comment: "live test cancellation" });
  }
  skip("calendar_respond_to_event", "needs an invitation from someone else");
  await step(srv, "calendar_list_rooms", {}, (v) => Array.isArray(arr(v)) || "no list");

  // ---- people ----
  const mgr = await srv.call("people_get_manager", {});
  record("people_get_manager", mgr.ok || /not ?found|Request_ResourceNotFound/i.test(mgr.text), mgr.ok ? mgr.value?.displayName ?? "" : `no manager set: ${mgr.text.slice(0, 80)}`);
  await step(srv, "people_list_direct_reports", {}, (v) => Array.isArray(arr(v)) || "no list");
  await step(srv, "people_get_profile", { user: self }, (v) => v.mail?.toLowerCase() === self.toLowerCase() || JSON.stringify(v).slice(0, 120));
  await step(srv, "people_get_presence", {}, (v) => !!v.availability || JSON.stringify(v).slice(0, 120));
  await step(srv, "people_get_presence", { users: [self] }, (v) => !!v[0]?.availability || JSON.stringify(v).slice(0, 160), "people_get_presence (by address)");
} catch (e) {
  console.log("ABORTED:", e.message);
} finally {
  await run.cleanup();
  run.save(dir);
  const failed = run.results.filter((r) => !r.pass);
  console.log(`\n${run.results.length - failed.length}/${run.results.length} passed`);
  srv.kill();
}
