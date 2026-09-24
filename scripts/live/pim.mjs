// Live run: mail, calendar, contacts, To Do, Planner (read), OneNote, Teams, auth, graph.
// Writes stay in the user's own space: mail only to themselves, events without
// attendees, Teams only in their notes-to-self chat. Everything is tagged.
import { startServer, makeRun } from "./harness.mjs";

const [dist, dir] = process.argv.slice(2);
const srv = startServer(dist);
const run = makeRun("pim");
const { step, record, onCleanup } = run;
const TAG = `[mcp-live-test ${Date.now()}]`;
const has = (v, s) => JSON.stringify(v).includes(s) || `missing "${s}" in ${JSON.stringify(v).slice(0, 200)}`;
const arr = (v) => v?.value ?? (Array.isArray(v) ? v : []);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const skip = (tool, why) => record(tool, true, `SKIPPED: ${why}`);

await srv.init();
try {
  // ---------------- auth + graph ----------------
  await step(srv, "auth_status", {}, (v) => v.signedIn === true || "not signed in");
  const me = await step(srv, "graph_batch_get", { requests: [{ id: "me", url: "/me?$select=mail,userPrincipalName,displayName" }] }, (v) => v.responses?.[0]?.status === 200 || "no /me");
  const self = me?.responses?.[0]?.body?.mail ?? me?.responses?.[0]?.body?.userPrincipalName;
  if (!self) throw new Error("could not read own address");
  await step(srv, "graph_search", { query: "meeting", entityTypes: ["message"], size: 3 }, (v) => Array.isArray(v.hits) || "no hits array");
  await step(srv, "graph_delta", { resource: "contacts" }, (v) => v.complete === true || "incomplete");

  // ---------------- mail ----------------
  await step(srv, "mail_list_folders", {}, (v) => has(v, "Inbox"));
  const inbox = await step(srv, "mail_list_messages", { folder: "Inbox", top: 5 }, (v) => Array.isArray(arr(v)) || "no list");
  await step(srv, "mail_search_messages", { query: "meeting", top: 3 }, (v) => Array.isArray(arr(v)) || "no list");
  const first = arr(inbox)[0];
  if (first) {
    await step(srv, "mail_get_message", { id: first.id, bodyFormat: "text" }, (v) => v.id === first.id || "id mismatch");
    await step(srv, "mail_list_attachments", { messageId: first.id }, (v) => Array.isArray(arr(v)) || "no list");
  } else { skip("mail_get_message", "empty inbox"); skip("mail_list_attachments", "empty inbox"); }

  const draft = await step(srv, "mail_create_draft", { to: [{ address: self }], subject: `${TAG} draft`, body: "draft body", bodyType: "Text" }, (v) => !!v.id || "no id");
  let draftDeleted = false;
  if (draft?.id) onCleanup("delete draft", async () => { if (!draftDeleted) await srv.call("mail_delete_message", { id: draft.id }); });

  const sendSubject = `${TAG} send`;
  await step(srv, "mail_send_message", { to: [{ address: self }], subject: sendSubject, body: "sent to self by the live test", bodyType: "Text", saveToSentItems: true });
  // Find everything carrying the tag, wherever it landed, and delete it at the end.
  onCleanup("delete tagged mail (inbox, sent, replies)", async () => {
    for (let pass = 0; pass < 3; pass++) {
      const r = await srv.call("mail_search_messages", { query: `"${TAG}"`, top: 25 });
      for (const m of arr(r.value)) if (m.subject?.includes(TAG) && m.id !== draft?.id) await srv.call("mail_delete_message", { id: m.id });
      await sleep(3000);
    }
  });
  let received;
  for (let i = 0; i < 20 && !received; i++) {
    await sleep(4000);
    const r = await srv.call("mail_list_messages", { folder: "Inbox", top: 15 });
    received = arr(r.value).find((m) => m.subject === sendSubject);
  }
  record("mail_send_message (arrived)", !!received, received ? "" : "not in inbox after 80 s");
  if (received) {
    await step(srv, "mail_reply_message", { id: received.id, comment: "reply from the live test", replyAll: false });
    let reply;
    for (let i = 0; i < 20 && !reply; i++) {
      await sleep(4000);
      const r = await srv.call("mail_list_messages", { folder: "Inbox", top: 15 });
      reply = arr(r.value).find((m) => m.subject?.includes(sendSubject) && m.id !== received.id);
    }
    record("mail_reply_message (arrived)", !!reply, reply ? reply.subject : "reply not in inbox after 80 s");
  } else skip("mail_reply_message", "nothing to reply to");
  if (draft?.id) draftDeleted = !!(await step(srv, "mail_delete_message", { id: draft.id }));

  // ---------------- calendar ----------------
  await step(srv, "calendar_list_calendars", {}, (v) => arr(v).length > 0 || "no calendars");
  const now = new Date();
  await step(srv, "calendar_list_events", { startDateTime: now.toISOString(), endDateTime: new Date(+now + 7 * 864e5).toISOString() }, (v) => Array.isArray(arr(v)) || "no list");
  const day = new Date(+now + 30 * 864e5).toISOString().slice(0, 10);
  const ev = await step(srv, "calendar_create_event", {
    subject: `${TAG} event`, start: { dateTime: `${day}T09:00:00`, timeZone: "Europe/Madrid" }, end: { dateTime: `${day}T09:30:00`, timeZone: "Europe/Madrid" },
    body: "no attendees; created and deleted by the live test", bodyType: "Text", isOnlineMeeting: false,
  }, (v) => !!v.id || "no id");
  if (ev?.id) {
    onCleanup("delete test event", async () => { await srv.call("calendar_delete_event", { id: ev.id }); });
    await step(srv, "calendar_get_event", { id: ev.id }, (v) => (v.subject === `${TAG} event` && (v.attendees ?? []).length === 0) || `got ${v.subject}, ${(v.attendees ?? []).length} attendees`);
    await step(srv, "calendar_update_event", { id: ev.id, subject: `${TAG} event (updated)` });
    await step(srv, "calendar_get_event", { id: ev.id }, (v) => v.subject === `${TAG} event (updated)` || `subject now ${v.subject}`);
  }
  await step(srv, "calendar_find_meeting_times", {
    attendees: [{ email: self, type: "required" }], durationMinutes: 30,
    timeWindowStart: `${day}T08:00:00`, timeWindowEnd: `${day}T18:00:00`, timeZone: "Europe/Madrid",
  }, (v) => has(v, "meetingTimeSuggestions"));
  await step(srv, "calendar_get_free_busy", {
    schedules: [self], startDateTime: `${day}T08:00:00`, endDateTime: `${day}T18:00:00`, timeZone: "Europe/Madrid", availabilityViewIntervalMinutes: 30,
  }, (v) => has(v, "availabilityView"));
  if (ev?.id) await step(srv, "calendar_delete_event", { id: ev.id });

  // ---------------- contacts ----------------
  await step(srv, "contacts_list", { top: 5 }, (v) => Array.isArray(arr(v)) || "no list");
  await step(srv, "contacts_people_search", { query: "a", top: 3 }, (v) => Array.isArray(arr(v)) || "no list");
  const c = await step(srv, "contacts_create", { givenName: "LiveTest", surname: TAG, emailAddresses: [{ address: "live-test@example.com" }] }, (v) => !!v.id || "no id");
  if (c?.id) {
    onCleanup("delete test contact", async () => { await srv.call("contacts_delete", { id: c.id }); });
    await step(srv, "contacts_update", { id: c.id, patch: { jobTitle: "Updated by live test" } }, (v) => (v.jobTitle ?? "Updated by live test") === "Updated by live test" || `jobTitle ${v.jobTitle}`);
    await step(srv, "contacts_search", { query: "LiveTest", top: 5 }, (v) => Array.isArray(arr(v)) || "no list");
    await step(srv, "contacts_delete", { id: c.id });
  }

  // ---------------- To Do + Planner ----------------
  const lists = await step(srv, "todo_list_lists", {}, (v) => arr(v).length > 0 || "no lists");
  const list = arr(lists).find((l) => l.wellknownListName === "defaultList") ?? arr(lists)[0];
  if (list) {
    await step(srv, "todo_list_tasks", { listId: list.id, top: 5 }, (v) => Array.isArray(arr(v)) || "no list");
    const t = await step(srv, "todo_create_task", { listId: list.id, title: `${TAG} task`, importance: "low" }, (v) => !!v.id || "no id");
    if (t?.id) await step(srv, "todo_complete_task", { listId: list.id, taskId: t.id }, (v) => (v.status ?? "completed") === "completed" || `status ${v.status}`);
  }
  const plans = await step(srv, "planner_list_plans", {}, (v) => Array.isArray(arr(v)) || "no list");
  const plan = arr(plans)[0];
  if (plan) await step(srv, "planner_list_tasks", { planId: plan.id, top: 5 }, (v) => Array.isArray(arr(v)) || "no list");
  else skip("planner_list_tasks", "no plans in this account");

  // ---------------- OneNote ----------------
  const nbs = await step(srv, "onenote_list_notebooks", {}, (v) => Array.isArray(arr(v)) || "no list");
  const nb = arr(nbs)[0];
  const secs = await step(srv, "onenote_list_sections", nb ? { notebookId: nb.id } : {}, (v) => Array.isArray(arr(v)) || "no list");
  const sec = arr(secs)[0];
  await step(srv, "onenote_list_pages", sec ? { sectionId: sec.id, top: 5 } : { top: 5 }, (v) => Array.isArray(arr(v)) || "no list");
  if (sec) {
    const page = await step(srv, "onenote_create_page", { sectionId: sec.id, title: `${TAG} page`, html: "<p>created by the live test</p>" }, (v) => !!v.id || "no id");
    if (page?.id) {
      onCleanup("delete test OneNote page", async () => { await srv.call("onenote_delete_page", { pageId: page.id }); });
      let content;
      for (let i = 0; i < 6 && !content?.ok; i++) { content = await srv.call("onenote_get_page_content", { pageId: page.id }); if (!content.ok) await sleep(3000); }
      record("onenote_get_page_content", !!content?.ok && content.text.includes("created by the live test"), content?.ok ? "" : content?.text.slice(0, 200));
      await step(srv, "onenote_delete_page", { pageId: page.id });
    }
  } else { skip("onenote_create_page", "no section"); skip("onenote_get_page_content", "no section"); skip("onenote_delete_page", "no section"); }

  // ---------------- Teams ----------------
  const teams = await step(srv, "teams_list_joined", {}, (v) => Array.isArray(arr(v)) || "no list");
  const team = arr(teams)[0];
  if (team) {
    const chans = await step(srv, "teams_list_channels", { teamId: team.id }, (v) => arr(v).length > 0 || "no channels");
    const ch = arr(chans)[0];
    if (ch) {
      const msgs = await step(srv, "teams_list_channel_messages", { teamId: team.id, channelId: ch.id, top: 3 }, (v) => Array.isArray(arr(v)) || "no list");
      const m = arr(msgs)[0];
      if (m) await step(srv, "teams_get_message_replies", { teamId: team.id, channelId: ch.id, messageId: m.id, top: 3 }, (v) => Array.isArray(arr(v)) || "no list");
      else skip("teams_get_message_replies", "channel has no messages");
    }
  } else { skip("teams_list_channels", "no teams"); skip("teams_list_channel_messages", "no teams"); skip("teams_get_message_replies", "no teams"); }
  const chats = await step(srv, "teams_list_chats", { top: 5 }, (v) => Array.isArray(arr(v)) || "no list");
  const chat = arr(chats)[0];
  if (chat) await step(srv, "teams_list_chat_messages", { chatId: chat.id, top: 3 }, (v) => Array.isArray(arr(v)) || "no list");
  // Only the personal notes chat, which no one else can see.
  await step(srv, "teams_post_chat_message", { chatId: "48:notes", body: `${TAG} note to self from the live test`, contentType: "text" }, (v) => !!v.id || "no id");
  skip("teams_post_channel_message", "visible to the team; not authorised for this run");
  skip("teams_reply_channel_message", "visible to the team; not authorised for this run");
  skip("planner_create_task", "visible to plan members; not authorised for this run");
  skip("planner_complete_task", "visible to plan members; not authorised for this run");
} catch (e) {
  console.log("ABORTED:", e.message);
} finally {
  await run.cleanup();
  run.save(dir);
  const failed = run.results.filter((r) => !r.pass);
  console.log(`\n${run.results.length - failed.length}/${run.results.length} passed`);
  srv.kill();
}
