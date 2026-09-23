import { createHash } from "node:crypto";
import { promises as fs } from "node:fs";
import type { AuthenticationRecord } from "@azure/identity";
import path from "node:path";
import os from "node:os";

export function defaultCachePath(): string {
  return path.join(os.homedir(), ".microsoft365-mcp", "tokencache.json");
}

/**
 * Name of the MSAL persistent cache the identity plugin keeps (under
 * %LOCALAPPDATA%\.IdentityService, ~/.IdentityService or the OS keychain).
 *
 * The plugin keys its store by this name alone, not by any path of ours, so a
 * fixed name made every server share one token cache whatever
 * MCP_TOKEN_CACHE_PATH said. The default path keeps the original name, so an
 * existing sign-in survives; any other path gets a cache of its own.
 */
export function persistentCacheName(cachePath: string): string {
  const base = "microsoft365-mcp";
  if (path.resolve(cachePath) === path.resolve(defaultCachePath())) return base;
  const digest = createHash("sha256").update(path.resolve(cachePath)).digest("hex").slice(0, 12);
  return `${base}-${digest}`;
}

export async function readCache(cachePath: string): Promise<string | undefined> {
  try {
    return await fs.readFile(cachePath, "utf8");
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code === "ENOENT") return undefined;
    throw err;
  }
}

export async function writeCache(cachePath: string, contents: string): Promise<void> {
  await fs.mkdir(path.dirname(cachePath), { recursive: true });
  await fs.writeFile(cachePath, contents, { mode: 0o600 });
  // Best-effort tighten on platforms that honor chmod.
  try {
    await fs.chmod(cachePath, 0o600);
  } catch {
    /* ignore */
  }
}

export function authRecordPath(cachePath: string): string {
  return path.join(path.dirname(cachePath), "authrecord.json");
}

/**
 * The AuthenticationRecord says which cached account to spend. Without it a fresh
 * process cannot pick an account out of the persistent cache silently, and falls
 * back to prompting — which is exactly the re-prompt-every-launch behaviour we
 * are trying to avoid.
 */
export async function readAuthRecord(cachePath: string): Promise<AuthenticationRecord | undefined> {
  const raw = await readCache(authRecordPath(cachePath));
  if (!raw) return undefined;
  try {
    return JSON.parse(raw) as AuthenticationRecord;
  } catch {
    return undefined;
  }
}

export async function writeAuthRecord(
  cachePath: string,
  record: AuthenticationRecord,
): Promise<void> {
  await writeCache(authRecordPath(cachePath), JSON.stringify(record));
}
