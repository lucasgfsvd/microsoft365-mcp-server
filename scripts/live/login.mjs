// Sign a token cache in once, so every live script can run against it.
// Settings come from the environment (MCP_TOKEN_CACHE_PATH, MCP_TOKEN_CACHE_KEY,
// MCP_ENABLE_SHARED_MAILBOXES, MCP_ENABLE_ADMIN_SCOPES...), as for the scripts.
// Usage: node scripts/live/login.mjs <dist>    Exits 0 once signed in.
import { startServer } from "./harness.mjs";

const [dist] = process.argv.slice(2);
const srv = startServer(dist);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
await srv.init();
let code = 1;
try {
  const signedIn = async () => (await srv.call("auth_status")).value?.signedIn === true;
  if (await signedIn()) {
    console.log("already signed in");
    code = 0;
  } else {
    const s = await srv.call("auth_sign_in");
    if (!s.value?.userCode) throw new Error(`no device code: ${s.text.slice(0, 200)}`);
    console.log(`\n>>> SIGN IN: open ${s.value.verificationUri} and enter ${s.value.userCode}\n`);
    for (let i = 0; i < 168; i++) {
      await sleep(5000);
      if (await signedIn()) {
        console.log("signed in");
        code = 0;
        break;
      }
    }
    if (code) console.log("not signed in after 14 minutes");
  }
} catch (e) {
  console.log("ABORTED:", e.message);
} finally {
  srv.kill();
  process.exit(code);
}
