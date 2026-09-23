import path from "node:path";
import type { ServerConfig } from "../types.js";
import { persistentCacheName } from "./tokenCache.js";

/**
 * The token store is owned by @azure/identity-cache-persistence, which offers no
 * way to clear it. These helpers reach the same store and remove this server's
 * tokens from it, so `--logout` actually logs out.
 *
 * Only credentials issued to *our* client id are removed. On macOS, and on Linux
 * with a keyring, the plugin keeps a single keychain item (fixed service and
 * account) shared by every app on the machine that uses it — deleting the item
 * outright could sign the user out of unrelated tools.
 */

/** Sections of MSAL's serialized cache that hold credentials, keyed by entry id. */
const CREDENTIAL_SECTIONS = ["AccessToken", "RefreshToken", "IdToken"] as const;

interface CacheEntry {
  client_id?: string;
  home_account_id?: string;
}
type SerializedCache = Record<string, Record<string, CacheEntry> | undefined>;

export interface PruneResult {
  json: string;
  /** Credentials removed. */
  removed: number;
  /** Nothing left in the store, from any app. */
  empty: boolean;
}

/**
 * Drop every credential issued to `clientId`, its app metadata, and the accounts
 * that were left with no credentials by doing so. Accounts other apps still hold
 * tokens for are kept.
 */
export function pruneClient(serialized: string, clientId: string): PruneResult {
  const cache = JSON.parse(serialized) as SerializedCache;
  const touchedAccounts = new Set<string>();
  let removed = 0;

  for (const section of CREDENTIAL_SECTIONS) {
    const entries = cache[section] ?? {};
    for (const [key, entry] of Object.entries(entries)) {
      if (entry.client_id !== clientId) continue;
      if (entry.home_account_id) touchedAccounts.add(entry.home_account_id);
      delete entries[key];
      removed++;
    }
  }
  for (const [key, entry] of Object.entries(cache.AppMetadata ?? {})) {
    if (entry.client_id === clientId) delete cache.AppMetadata![key];
  }

  const stillHeld = new Set(
    CREDENTIAL_SECTIONS.flatMap((s) => Object.values(cache[s] ?? {}).map((e) => e.home_account_id)),
  );
  for (const [key, entry] of Object.entries(cache.Account ?? {})) {
    const id = entry.home_account_id;
    if (id && touchedAccounts.has(id) && !stillHeld.has(id)) delete cache.Account![key];
  }

  const empty = Object.values(cache).every((section) => !section || Object.keys(section).length === 0);
  return { json: JSON.stringify(cache), removed, empty };
}

/** Structural slice of msal-node-extensions' persistence we rely on. */
interface Persistence {
  delete(): Promise<boolean>;
  getFilePath(): string;
}
type Extensions = typeof import("@azure/msal-node-extensions");

/**
 * Open the store exactly where the cache plugin puts it (mirrors
 * identity-cache-persistence's `msalPersistencePlatforms`): a DPAPI file on
 * Windows, the keychain on macOS, libsecret on Linux, and a plain file where the
 * keychain is unreadable (the server allows unencrypted fallback).
 */
async function openPersistence(ext: Extensions, name: string): Promise<Persistence> {
  const localAppData = process.env.APPDATA?.replace(/(.Roaming)*$/, "\\Local") ?? process.env.HOME ?? "";
  const file = path.join(localAppData, ".IdentityService", name);
  const service = "Microsoft.Developer.IdentityService";
  const account = "MSALCache";

  if (process.platform === "win32") {
    return ext.FilePersistenceWithDataProtection.create(file, ext.DataProtectionScope.CurrentUser);
  }
  const keyed =
    process.platform === "darwin"
      ? () => ext.KeychainPersistence.create(file, service, account)
      : () => ext.LibSecretPersistence.create(file, service, account);
  try {
    const p = await keyed();
    await p.load();
    return p;
  } catch {
    return ext.FilePersistence.create(file);
  }
}

export type ClearOutcome =
  | { status: "cleared"; removed: number; deletedStore: boolean; location: string }
  | { status: "nothing-stored"; location: string }
  | { status: "unavailable"; reason: string };

/**
 * Remove this server's tokens from the persistent store.
 *
 * Runs inside the plugin's own lock (via PersistenceCachePlugin's before/after
 * hooks), so a server using the store concurrently cannot interleave a write.
 */
export async function clearStoredTokens(
  config: Pick<ServerConfig, "clientId" | "tokenCachePath">,
  load: () => Promise<Extensions> = () => import("@azure/msal-node-extensions"),
): Promise<ClearOutcome> {
  let ext: Extensions;
  try {
    ext = await load();
  } catch (err) {
    // Same failure that stops the server persisting anything (e.g. no libsecret):
    // tokens were only ever held in memory.
    return { status: "unavailable", reason: err instanceof Error ? err.message : String(err) };
  }

  const persistence = await openPersistence(ext, persistentCacheName(config.tokenCachePath));
  const location = persistence.getFilePath();
  const plugin = new ext.PersistenceCachePlugin(
    persistence as ConstructorParameters<Extensions["PersistenceCachePlugin"]>[0],
  );

  let stored: string | undefined;
  let next = "";
  // The hooks take MSAL's TokenCacheContext; only these members are used.
  const ctx = {
    cacheHasChanged: true,
    tokenCache: {
      deserialize: (s: string) => {
        stored = s;
      },
      serialize: () => next,
    },
  };
  const hookCtx = ctx as unknown as Parameters<InstanceType<Extensions["PersistenceCachePlugin"]>["beforeCacheAccess"]>[0];

  await plugin.beforeCacheAccess(hookCtx); // takes the lock and loads
  let outcome: ClearOutcome;
  try {
    if (!stored) {
      ctx.cacheHasChanged = false;
      // Opening the store creates an empty file; don't leave that behind.
      await persistence.delete();
      outcome = { status: "nothing-stored", location };
    } else {
      const pruned = pruneClient(stored, config.clientId);
      next = pruned.json;
      ctx.cacheHasChanged = pruned.removed > 0 && !pruned.empty;
      if (pruned.empty) await persistence.delete();
      outcome = { status: "cleared", removed: pruned.removed, deletedStore: pruned.empty, location };
    }
  } catch (err) {
    ctx.cacheHasChanged = false;
    throw err;
  } finally {
    await plugin.afterCacheAccess(hookCtx); // writes if changed, releases the lock
  }
  return outcome;
}
