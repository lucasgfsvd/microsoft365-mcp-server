import { promises as fs } from "node:fs";

export interface LockOptions {
  /** How often to try again while another holder has the lock. */
  delayMs?: number;
  /** Give up waiting after this long. */
  timeoutMs?: number;
  /** A lock older than this was left by a process that died holding it. */
  staleMs?: number;
}

/**
 * Take a lock shared between processes: a file created exclusively, so only one
 * holder exists at a time. Resolves to a function that releases it.
 *
 * Waiting has a limit. Past it the caller proceeds unlocked (and is told so),
 * since a token write racing another is less harmful than failing a sign-in.
 */
export async function acquireLock(
  lockPath: string,
  { delayMs = 50, timeoutMs = 5_000, staleMs = 10_000 }: LockOptions = {},
): Promise<{ release: () => Promise<void>; locked: boolean }> {
  const deadline = Date.now() + timeoutMs;
  for (;;) {
    try {
      const handle = await fs.open(lockPath, "wx");
      await handle.writeFile(String(process.pid)).finally(() => handle.close());
      return { locked: true, release: () => fs.unlink(lockPath).catch(() => undefined) };
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code !== "EEXIST") throw err;
    }
    const age = await fs.stat(lockPath).then(
      (s) => Date.now() - s.mtimeMs,
      () => 0, // released between our attempt and this check: just try again
    );
    if (age > staleMs) {
      await fs.unlink(lockPath).catch(() => undefined);
      continue;
    }
    if (Date.now() >= deadline) return { locked: false, release: async () => undefined };
    await new Promise((r) => setTimeout(r, delayMs));
  }
}
