// Live-test harness: drives a separate server over MCP stdio, records a verdict
// per tool, and always runs registered cleanups (newest first).
import { spawn } from "node:child_process";
import { writeFileSync } from "node:fs";

export function startServer(dist, env = {}) {
  const srv = spawn(process.execPath, [dist], {
    stdio: ["pipe", "pipe", "pipe"],
    env: { ...process.env, MCP_ENABLE_WRITES: "1", ...env },
  });
  let stderr = "";
  srv.stderr.on("data", (d) => (stderr += d));
  srv.stdin.on("error", () => {});
  srv.stdout.setEncoding("utf8");
  let buf = "", id = 0;
  const waiting = new Map();
  srv.stdout.on("data", (d) => {
    buf += d;
    let i;
    while ((i = buf.indexOf("\n")) >= 0) {
      const line = buf.slice(0, i);
      buf = buf.slice(i + 1);
      if (!line.trim()) continue;
      const m = JSON.parse(line);
      if (m.id && waiting.has(m.id)) { waiting.get(m.id)(m); waiting.delete(m.id); }
    }
  });
  const send = (method, params) =>
    new Promise((res) => { const n = ++id; waiting.set(n, res); srv.stdin.write(JSON.stringify({ jsonrpc: "2.0", id: n, method, params }) + "\n"); });

  /** Call a tool; returns { ok, value, text }. Never throws. */
  const call = async (name, args = {}) => {
    const r = await send("tools/call", { name, arguments: args });
    if (r.error) return { ok: false, text: `protocol error: ${r.error.message}` };
    const text = r.result?.content?.[0]?.text ?? "";
    if (r.result?.isError) return { ok: false, text };
    let value; try { value = JSON.parse(text); } catch { value = text; }
    return { ok: true, value, text };
  };
  const init = async () => {
    await send("initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "live", version: "1" } });
    srv.stdin.write(JSON.stringify({ jsonrpc: "2.0", method: "notifications/initialized" }) + "\n");
  };
  return { call, init, kill: () => srv.kill(), stderr: () => stderr };
}

export function makeRun(label) {
  const results = [];
  const cleanups = [];
  const record = (tool, pass, note = "") => {
    results.push({ tool, pass, note });
    console.log(`${pass ? "PASS" : "FAIL"}  ${tool.padEnd(32)} ${note}`.slice(0, 300));
  };
  /** Run a tool, check the result with `verify` (return true or a failure string). */
  const step = async (srv, tool, args, verify = () => true) => {
    const r = await srv.call(tool, args);
    if (!r.ok) { record(tool, false, r.text.slice(0, 250)); return undefined; }
    let v;
    try { v = verify(r.value); } catch (e) { v = `verify threw: ${e.message}`; }
    record(tool, v === true, v === true ? "" : String(v));
    return r.value;
  };
  const onCleanup = (desc, fn) => cleanups.push({ desc, fn });
  const cleanup = async () => {
    for (const c of cleanups.reverse()) {
      try { await c.fn(); console.log(`cleanup ok    ${c.desc}`); } catch (e) { console.log(`CLEANUP FAILED ${c.desc}: ${e.message}`); }
    }
  };
  const save = (dir) => writeFileSync(`${dir}/results-${label}.json`, JSON.stringify(results, null, 1));
  return { step, record, onCleanup, cleanup, results, save };
}
