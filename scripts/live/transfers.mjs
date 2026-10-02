// Live run: progress and cancellation of large transfers, inside one throwaway
// OneDrive folder. Usage: node scripts/live/transfers.mjs <dist> <scratch>
import { createHash, randomBytes } from "node:crypto";
import { mkdirSync, writeFileSync } from "node:fs";
import { startServer, makeRun } from "./harness.mjs";

const [dist, dir] = process.argv.slice(2);
const upDir = `${dir}/up`;
const dlDir = `${dir}/dl`;
mkdirSync(upDir, { recursive: true });
mkdirSync(dlDir, { recursive: true });
const srv = startServer(dist, { MCP_UPLOAD_DIR: upDir, MCP_DOWNLOAD_DIR: dlDir });
const run = makeRun("transfers");
const { step, record, onCleanup } = run;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const folderName = `mcp-live-test-${Date.now()}`;
const root = `/${folderName}`;

/** Progress must rise, and end at the total. */
const rising = (p, total) => p.length > 0 && p.every((x, i) => i === 0 || x.progress > p[i - 1].progress) && p.at(-1).progress === total && p.at(-1).total === total;

await srv.init();
try {
  const folder = await step(srv, "files_create_folder", { parentPath: "/", name: folderName }, (v) => v.name === folderName || "wrong name");
  if (!folder) throw new Error("no sandbox folder");
  onCleanup(`delete sandbox folder ${folderName}`, async () => { const r = await srv.call("files_delete", { itemId: folder.id }); if (!r.ok) throw new Error(r.text); });

  // A 24 MB upload, with progress: three chunks of 10 MiB at most.
  const big = randomBytes(24 * 1024 * 1024);
  writeFileSync(`${upDir}/progress.bin`, big);
  const up = srv.startCall("files_upload", { parentPath: root, localPath: "progress.bin" }, "up");
  const upped = await up.result;
  const upProgress = srv.progressFor("up");
  record("files_upload (progress)", upped.ok && upProgress.length >= 2 && rising(upProgress, big.length),
    upped.ok ? upProgress.map((p) => p.message).join(" | ").slice(0, 250) : upped.text.slice(0, 200));

  if (upped.ok) {
    const down = srv.startCall("files_download", { itemId: upped.value.id, saveToDisk: true }, "down");
    const got = await down.result;
    const downProgress = srv.progressFor("down");
    const same = got.ok && got.value.sha256 === createHash("sha256").update(big).digest("hex");
    record("files_download (progress)", same && rising(downProgress, big.length), got.ok ? `${downProgress.length} updates` : got.text.slice(0, 200));
  }

  // Cancelled after the first chunk lands: no answer, no file, the server carries on.
  writeFileSync(`${upDir}/cancel.bin`, randomBytes(48 * 1024 * 1024));
  const doomed = srv.startCall("files_upload", { parentPath: root, localPath: "cancel.bin" }, "cancel");
  for (let i = 0; i < 120 && srv.progressFor("cancel").length === 0; i++) await sleep(500);
  srv.cancel(doomed.id);
  const answered = await Promise.race([doomed.result.then(() => true), sleep(15_000).then(() => false)]);
  record("files_upload (cancelled: no answer)", !answered && srv.progressFor("cancel").length > 0, answered ? "the server answered a cancelled call" : "");
  await sleep(3000);
  await step(srv, "files_list_children", { path: root }, (v) => !JSON.stringify(v).includes("cancel.bin") || "cancelled upload left a file", "files_upload (cancelled: no file)");
  await step(srv, "auth_status", {}, (v) => v.signedIn === true || "server not answering after a cancel", "server still answering after a cancel");
} catch (e) {
  console.log("ABORTED:", e.message);
} finally {
  await run.cleanup();
  run.save(dir);
  const failed = run.results.filter((r) => !r.pass);
  console.log(`\n${run.results.length - failed.length}/${run.results.length} passed`);
  srv.kill();
}
