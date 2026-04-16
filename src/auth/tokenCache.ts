import { promises as fs } from "node:fs";
import path from "node:path";
import os from "node:os";

export function defaultCachePath(): string {
  return path.join(os.homedir(), ".microsoft365-mcp", "tokencache.json");
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
