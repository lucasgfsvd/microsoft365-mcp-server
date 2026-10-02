// Live run: Teams channel posts and Planner tasks (created, completed, deleted), inside a sandbox team and plan
// that nobody else is in. Usage:
//   node scripts/live/teams-planner.mjs <dist> <scratch> <teamId> <planId>
import { startServer, makeRun } from "./harness.mjs";

const [dist, dir, teamId, planId] = process.argv.slice(2);
if (!teamId || !planId) throw new Error("pass the sandbox team id and plan id");
const srv = startServer(dist);
const run = makeRun("teams-planner");
const { step, record } = run;
const TAG = `[mcp-live-test ${Date.now()}]`;
const arr = (v) => v?.value ?? (Array.isArray(v) ? v : []);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

await srv.init();
try {
  // ---------------- Teams channel ----------------
  const chans = await step(srv, "teams_list_channels", { teamId }, (v) => arr(v).length > 0 || "no channels");
  const channel = arr(chans).find((c) => c.displayName === "tests") ?? arr(chans)[0];
  if (!channel) throw new Error("no channel in the sandbox team");

  const post = await step(srv, "teams_post_channel_message", {
    teamId, channelId: channel.id, body: `${TAG} channel post from the live test`, contentType: "text",
  }, (v) => !!v.id || "no id");
  if (post?.id) {
    let listed;
    for (let i = 0; i < 6 && !listed; i++) {
      const r = await srv.call("teams_list_channel_messages", { teamId, channelId: channel.id, top: 10 });
      listed = r.ok && arr(r.value).some((m) => m.id === post.id);
      if (!listed) await sleep(3000);
    }
    record("teams_list_channel_messages (sees post)", !!listed, listed ? "" : "post not listed");

    const reply = await step(srv, "teams_reply_channel_message", {
      teamId, channelId: channel.id, messageId: post.id, body: `${TAG} reply from the live test`, contentType: "text",
    }, (v) => !!v.id || "no id");
    if (reply?.id) {
      let seen;
      for (let i = 0; i < 6 && !seen; i++) {
        const r = await srv.call("teams_get_message_replies", { teamId, channelId: channel.id, messageId: post.id });
        seen = r.ok && arr(r.value).some((m) => m.id === reply.id);
        if (!seen) await sleep(3000);
      }
      record("teams_get_message_replies (sees reply)", !!seen, seen ? "" : "reply not listed");
    }
  }

  // ---------------- Planner ----------------
  await step(srv, "planner_list_plans", {}, (v) => arr(v).some((p) => p.id === planId) || "sandbox plan not listed");
  const task = await step(srv, "planner_create_task", { planId, title: `${TAG} task` }, (v) => (!!v.id && !!v["@odata.etag"]) || `got ${JSON.stringify(v).slice(0, 120)}`);
  if (task?.id) {
    // Planner is eventually consistent: a new task can take a few seconds to list.
    let listed;
    for (let i = 0; i < 6 && !listed; i++) {
      const r = await srv.call("planner_list_tasks", { planId });
      listed = r.ok && arr(r.value).some((t) => t.id === task.id);
      if (!listed) await sleep(2000);
    }
    record("planner_list_tasks", !!listed, listed ? "" : "new task not listed after 12 s");
    // Planner can change a new task's etag moments after creation; use the current one.
    const current = await srv.call("planner_list_tasks", { planId });
    const etag = arr(current.value).find((t) => t.id === task.id)?.["@odata.etag"] ?? task["@odata.etag"];
    await step(srv, "planner_complete_task", { taskId: task.id, etag });
    let done;
    for (let i = 0; i < 5 && !done; i++) {
      const r = await srv.call("planner_list_tasks", { planId });
      done = arr(r.value).find((t) => t.id === task.id)?.percentComplete === 100;
      if (!done) await sleep(2000);
    }
    record("planner_complete_task (is complete)", !!done, done ? "" : "task not at 100%");
    // No etag: the tool reads the current one, which completing the task just changed.
    await step(srv, "planner_delete_task", { taskId: task.id });
    let gone;
    for (let i = 0; i < 5 && !gone; i++) {
      const r = await srv.call("planner_list_tasks", { planId });
      gone = r.ok && !arr(r.value).some((t) => t.id === task.id);
      if (!gone) await sleep(2000);
    }
    record("planner_delete_task (is gone)", !!gone, gone ? "" : "task still listed");
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
