import { promises as fs } from "node:fs";
import path from "node:path";
import { allTools } from "./tools/index.js";
import { isToolAllowed } from "./util/writeGuard.js";
import { SURFACES, type ServerConfig } from "./types.js";
import { authRecordPath } from "./auth/tokenCache.js";

/**
 * Print the tool catalogue grouped by surface, marking which tools are
 * currently visible (writes enabled, not in disabled-tools) and which scopes
 * each one requires. Writes to stdout — this command never starts the MCP
 * server, so stdout is free for human-readable output.
 */
export function runListTools(config: ServerConfig): void {
  const tools = allTools();
  const lines: string[] = [];

  for (const surface of SURFACES) {
    const surfaceTools = tools.filter((t) => t.surface === surface);
    if (!surfaceTools.length) continue;
    lines.push(`# ${surface} (${surfaceTools.length})`);
    for (const t of surfaceTools) {
      const visible = isToolAllowed(t, config);
      const marker = visible ? " " : "·"; // · = hidden under current config
      const tag = t.mutating ? "[write]" : "[read] ";
      const scopes = (t.requiredScopes ?? []).join(", ") || "—";
      lines.push(`  ${marker} ${tag} ${t.name.padEnd(36)} ${scopes}`);
    }
    lines.push("");
  }

  const total = tools.length;
  const visibleCount = tools.filter((t) => isToolAllowed(t, config)).length;
  lines.push(
    `${total} tools (${visibleCount} visible under current config; · = filtered out by writeGuard or MCP_DISABLED_TOOLS)`,
  );

  process.stdout.write(lines.join("\n") + "\n");
}

/**
 * Delete the cached OAuth token file. Best-effort: ENOENT is treated as
 * already-logged-out. Note that on platforms with `keytar`, the OS keychain
 * may also hold a copy of the cache — that gets overwritten on the next
 * sign-in, but if you're trying to fully revoke access, do it from
 * https://myapps.microsoft.com as well.
 */
export async function runLogout(config: ServerConfig): Promise<void> {
  // Both files matter. The record names which account to reuse, so leaving it
  // behind means the next start believes it is still signed in to an account
  // whose token we just deleted.
  const targets = [config.tokenCachePath, authRecordPath(config.tokenCachePath)];
  let removed = 0;

  for (const target of targets) {
    try {
      await fs.unlink(target);
      process.stdout.write(`Removed ${target}\n`);
      removed++;
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code !== "ENOENT") throw err;
    }
  }

  if (removed === 0) {
    process.stdout.write(
      `Nothing to remove under ${path.dirname(config.tokenCachePath)} - already logged out.\n`,
    );
  }
  process.stdout.write(
    "Note: where a keychain is in use the OS store may still hold a copy until the " +
      "next sign-in overwrites it. To revoke access outright, use https://myapps.microsoft.com.\n",
  );
}
