// Live run: a sign-in that other servers see at once, and that outlives the
// process, kept in a token cache file encrypted with MCP_TOKEN_CACHE_KEY.
// Needs a person to enter one device code. Usage:
//   node scripts/live/signin.mjs <dist> <scratch> [docker-image]
// With an image, also checks that a container (no OS keyring) starts signed in
// from the same file. Everything lives in a fresh folder under <scratch>, which
// is signed out with --logout and removed at the end.
import { execFileSync, spawnSync } from "node:child_process";
import { createHash, randomBytes } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, rmSync } from "node:fs";
import path from "node:path";
import { startServer, makeRun } from "./harness.mjs";

const [dist, scratch, image] = process.argv.slice(2);
const dir = path.resolve(scratch, `signin-${Date.now()}`);
mkdirSync(dir, { recursive: true });
const file = path.join(dir, "tokencache.json");
const key = randomBytes(32).toString("base64");
const env = { MCP_TOKEN_CACHE_PATH: file, MCP_TOKEN_CACHE_KEY: key };
const run = makeRun("signin");
const { step, record } = run;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const sha = () => createHash("sha256").update(readFileSync(file)).digest("hex");
const signedIn = async (srv) => (await srv.call("auth_status")).value?.signedIn === true;
const callMe = (srv, label) =>
  step(srv, "graph_batch_get", { requests: [{ id: "me", url: "/me?$select=id" }] }, (v) => v.responses?.[0]?.status === 200 || "no /me", label);
const started = [];
const start = async (extra = {}, command) => {
  const srv = startServer(dist, { ...env, ...extra }, command);
  started.push(srv);
  await srv.init();
  return srv;
};

try {
  const a = await start();
  const b = await start();
  record("fresh cache: B starts signed out", !(await signedIn(b)));

  const s = await a.call("auth_sign_in");
  if (!s.ok || !s.value?.userCode) throw new Error(`no device code: ${s.text.slice(0, 200)}`);
  console.log(`\n>>> SIGN IN: open ${s.value.verificationUri} and enter ${s.value.userCode}\n`);
  let done = false;
  for (let i = 0; i < 168 && !done; i++) { await sleep(5000); done = await signedIn(a); }
  record("A: sign-in completed", done, done ? "" : "not signed in after 14 minutes");
  if (!done) throw new Error("no sign-in, nothing more to check");

  // B was running all along, with no account: it should notice A's sign-in.
  let seen = false;
  for (let i = 0; i < 10 && !seen; i++) { seen = await signedIn(b); if (!seen) await sleep(1000); }
  record("B: sees A's sign-in without a restart", seen);
  await callMe(b, "B: Graph call");

  const raw = readFileSync(file, "utf8");
  record("cache file is encrypted", raw.includes('"format":"microsoft365-mcp/encrypted-token-cache"') && !/refresh|secret|eyJ0/i.test(raw), raw.slice(0, 80));

  for (const srv of started.splice(0)) srv.kill();
  const c = await start();
  record("C: signed in after a restart", await signedIn(c));
  await callMe(c, "C: Graph call");
  c.kill();

  const before = sha();
  const d = await start({ MCP_TOKEN_CACHE_KEY: randomBytes(32).toString("base64") });
  record("D, wrong key: signed out, file untouched", !(await signedIn(d)) && sha() === before);
  d.kill();

  if (image) {
    // The container has no keyring at all; the file and the key are all it has.
    const name = `mcp-live-signin-${Date.now()}`;
    const docker = ["docker", ["run", "--rm", "-i", "--name", name, "-e", "MCP_TOKEN_CACHE_KEY", "-e", "MCP_ENABLE_WRITES=1", "-v", `${dir}:/data`, image]];
    try {
      const e = await start({}, docker);
      record("container: signed in from the file", await signedIn(e));
      await callMe(e, "container: Graph call");
      e.kill();
    } finally {
      spawnSync("docker", ["rm", "-f", name], { stdio: "ignore" });
    }
  }

  const out = execFileSync(process.execPath, [dist, "--logout"], { env: { ...process.env, ...env }, encoding: "utf8" });
  record("--logout removes the encrypted file", !existsSync(file) && !existsSync(path.join(dir, "authrecord.json")), out.trim().split("\n")[0]);
  const f = await start();
  record("F: signed out after --logout", !(await signedIn(f)));
} catch (e) {
  console.log("ABORTED:", e.message);
} finally {
  for (const srv of started) srv.kill();
  await sleep(500);
  run.save(scratch);
  rmSync(dir, { recursive: true, force: true });
  const failed = run.results.filter((r) => !r.pass);
  console.log(`\n${run.results.length - failed.length}/${run.results.length} passed`);
}
