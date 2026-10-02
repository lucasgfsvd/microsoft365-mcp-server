import { useIdentityPlugin } from "@azure/identity";
import { logger } from "../util/logger.js";
import type { ServerConfig } from "../types.js";
import { EncryptedFile, UnreadableFileError } from "./encryptedFile.js";
import { acquireLock } from "./fileLock.js";
import { pruneClient, type ClearOutcome } from "./tokenStore.js";

/**
 * A token cache in one encrypted file, for hosts with no OS keyring: the
 * distroless image above all, where a device-code sign-in otherwise lasts only
 * until the container restarts.
 *
 * It deliberately needs nothing native. The OS-store plugin cannot even be
 * imported there (keytar dlopens libsecret at load), so this talks to MSAL's
 * cache-plugin interface directly: read and decrypt before each cache access,
 * encrypt and write after one that changed something, all under a lock file so
 * servers sharing the file take turns.
 */

/** The slice of MSAL's TokenCacheContext a cache plugin uses. */
export interface CacheContext {
  cacheHasChanged: boolean;
  tokenCache: { serialize(): string; deserialize(cache: string): void };
}

export interface CachePlugin {
  beforeCacheAccess(ctx: CacheContext): Promise<void>;
  afterCacheAccess(ctx: CacheContext): Promise<void>;
}

const lockPath = (file: string) => `${file}.lock`;

export function encryptedCachePlugin(store: EncryptedFile): CachePlugin {
  // MSAL hands the same context to both hooks of one access, so it pairs them.
  const held = new WeakMap<CacheContext, () => Promise<void>>();
  let warned = false;
  return {
    async beforeCacheAccess(ctx) {
      const lock = await acquireLock(lockPath(store.file));
      held.set(ctx, lock.release);
      if (!lock.locked) logger.warn({ file: store.file }, "auth: token cache lock not released in time; reading without it");
      try {
        const text = await store.read();
        if (text) ctx.tokenCache.deserialize(text);
      } catch (err) {
        // Unreadable means signed out here, not broken: a new sign-in writes a fresh file.
        if (!warned) logger.warn({ err: err instanceof Error ? err.message : String(err) }, "auth: ignoring the token cache file");
        warned = true;
      }
    },
    async afterCacheAccess(ctx) {
      try {
        if (ctx.cacheHasChanged) await store.write(ctx.tokenCache.serialize());
      } finally {
        await held.get(ctx)?.();
        held.delete(ctx);
      }
    },
  };
}

let registered: string | undefined;

/**
 * Make every credential that asks for persistence use the encrypted file.
 * @azure/identity has one persistence provider per process, which this replaces.
 */
export function useEncryptedFileCache(config: Pick<ServerConfig, "tokenCachePath" | "tokenCacheKey">): void {
  if (!config.tokenCacheKey) throw new Error("useEncryptedFileCache needs MCP_TOKEN_CACHE_KEY.");
  if (registered === config.tokenCachePath) return;
  const store = new EncryptedFile(config.tokenCachePath, config.tokenCacheKey);
  useIdentityPlugin((context) => {
    const control = (context as { cachePluginControl: { setPersistence(factory: () => Promise<CachePlugin>): void } }).cachePluginControl;
    control.setPersistence(async () => encryptedCachePlugin(store));
  });
  registered = config.tokenCachePath;
  logger.info({ file: config.tokenCachePath }, "auth: token cache is the encrypted file (MCP_TOKEN_CACHE_KEY)");
}

/** `--logout` for the encrypted file: this app's tokens out, other apps' kept. */
export async function clearEncryptedTokens(config: Pick<ServerConfig, "clientId" | "tokenCachePath" | "tokenCacheKey">): Promise<ClearOutcome> {
  const store = new EncryptedFile(config.tokenCachePath, config.tokenCacheKey ?? "");
  const lock = await acquireLock(lockPath(store.file));
  try {
    const text = await store.read();
    if (!text) return { status: "nothing-stored", location: store.file };
    const pruned = pruneClient(text, config.clientId);
    if (pruned.empty) await store.remove();
    else if (pruned.removed > 0) await store.write(pruned.json);
    return { status: "cleared", removed: pruned.removed, deletedStore: pruned.empty, location: store.file };
  } catch (err) {
    if (err instanceof UnreadableFileError) return { status: "unavailable", reason: `${err.message}; delete the file to sign out` };
    throw err;
  } finally {
    await lock.release();
  }
}
