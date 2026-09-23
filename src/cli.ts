import { promises as fs } from "node:fs";
import { allTools } from "./tools/index.js";
import { isToolAllowed } from "./util/writeGuard.js";
import { SURFACES, type ServerConfig } from "./types.js";
import { authRecordPath } from "./auth/tokenCache.js";
import { clearStoredTokens } from "./auth/tokenStore.js";

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
 * Sign out: remove this server's tokens from the persistent store, and the
 * record naming which account to reuse. Either alone is not enough — the record
 * without tokens makes a restart believe it is signed in, and tokens without the
 * record leave a usable refresh token on disk.
 *
 * Revoking the grant itself is separate: https://myapps.microsoft.com.
 */
export async function runLogout(
  config: ServerConfig,
  clear: typeof clearStoredTokens = clearStoredTokens,
): Promise<void> {
  const out = (line: string) => process.stdout.write(line + "\n");

  const store = await clear(config);
  switch (store.status) {
    case "cleared":
      out(
        store.deletedStore
          ? `Removed the token store ${store.location}`
          : `Removed ${store.removed} token(s) for this app from ${store.location} (other apps' tokens kept)`,
      );
      break;
    case "nothing-stored":
      out(`No tokens stored in ${store.location}`);
      break;
    case "unavailable":
      out(`No persistent token store on this system (${store.reason}); nothing to clear there.`);
      break;
  }

  const record = authRecordPath(config.tokenCachePath);
  try {
    await fs.unlink(record);
    out(`Removed ${record}`);
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code !== "ENOENT") throw err;
    out(`No authentication record at ${record}`);
  }
  out("To revoke this app's access outright, use https://myapps.microsoft.com.");
}
