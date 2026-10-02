import { describe, it, expect, beforeEach, afterEach, vi } from "vitest";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { DeviceCodeCredential } from "@azure/identity";
import type { ServerConfig } from "../src/types.js";
import { EncryptedFile, UnreadableFileError } from "../src/auth/encryptedFile.js";
import { clearEncryptedTokens, encryptedCachePlugin, type CacheContext } from "../src/auth/encryptedCache.js";
import { acquireLock } from "../src/auth/fileLock.js";
import { buildCredential } from "../src/auth/index.js";

const KEY = "correct horse battery staple";
const SECRET_CACHE = JSON.stringify({ RefreshToken: { rt: { client_id: "cid", home_account_id: "h", secret: "0.AAAA-refresh" } } });

let dir: string;
let file: string;
beforeEach(async () => {
  dir = await fs.mkdtemp(path.join(os.tmpdir(), "m365-cache-"));
  file = path.join(dir, "tokencache.json");
});
afterEach(async () => {
  vi.restoreAllMocks();
  await fs.rm(dir, { recursive: true, force: true });
});

/** A stand-in for MSAL's TokenCacheContext. */
function context(changed: boolean, serialized = SECRET_CACHE) {
  const seen: string[] = [];
  const ctx: CacheContext = { cacheHasChanged: changed, tokenCache: { serialize: () => serialized, deserialize: (s) => void seen.push(s) } };
  return { ctx, seen };
}

describe("EncryptedFile", () => {
  it("round-trips, and keeps nothing readable on disk", async () => {
    await new EncryptedFile(file, KEY).write(SECRET_CACHE);
    const raw = await fs.readFile(file, "utf8");
    expect(raw).not.toContain("refresh");
    expect(raw).not.toContain("client_id");
    expect(await new EncryptedFile(file, KEY).read()).toBe(SECRET_CACHE);
  });

  it("refuses a wrong key, a tampered file and a file that is not one", async () => {
    await new EncryptedFile(file, KEY).write(SECRET_CACHE);
    await expect(new EncryptedFile(file, "a different key entirely").read()).rejects.toThrow(/does not match/);

    const env = JSON.parse(await fs.readFile(file, "utf8"));
    env.data = Buffer.from("tampered").toString("base64");
    await fs.writeFile(file, JSON.stringify(env));
    await expect(new EncryptedFile(file, KEY).read()).rejects.toBeInstanceOf(UnreadableFileError);

    await fs.writeFile(file, '{"Account":{}}');
    await expect(new EncryptedFile(file, KEY).read()).rejects.toThrow(/not an encrypted token cache/);
  });

  it("reads nothing when there is no file, and removes only what exists", async () => {
    const f = new EncryptedFile(file, KEY);
    expect(await f.read()).toBeUndefined();
    expect(await f.remove()).toBe(false);
    await f.write("x");
    expect(await f.remove()).toBe(true);
  });

  it("uses a fresh IV on every write", async () => {
    const f = new EncryptedFile(file, KEY);
    await f.write(SECRET_CACHE);
    const first = JSON.parse(await fs.readFile(file, "utf8"));
    await f.write(SECRET_CACHE);
    const second = JSON.parse(await fs.readFile(file, "utf8"));
    expect(second.salt).toBe(first.salt);
    expect(second.iv).not.toBe(first.iv);
    expect(second.data).not.toBe(first.data);
  });
});

describe("encryptedCachePlugin", () => {
  it("loads the cache before an access and saves it after one that changed it", async () => {
    const store = new EncryptedFile(file, KEY);
    const plugin = encryptedCachePlugin(store);
    const first = context(true);
    await plugin.beforeCacheAccess(first.ctx);
    expect(first.seen).toEqual([]); // nothing stored yet
    await plugin.afterCacheAccess(first.ctx);

    const second = context(false, "never written");
    await plugin.beforeCacheAccess(second.ctx);
    expect(second.seen).toEqual([SECRET_CACHE]);
    await plugin.afterCacheAccess(second.ctx);
    expect(await store.read()).toBe(SECRET_CACHE);
  });

  it("holds the lock for the whole access, so servers sharing the file take turns", async () => {
    const plugin = encryptedCachePlugin(new EncryptedFile(file, KEY));
    const a = context(true);
    await plugin.beforeCacheAccess(a.ctx);
    await expect(fs.stat(`${file}.lock`)).resolves.toBeTruthy();

    const order: string[] = [];
    const b = context(false);
    const waiting = plugin.beforeCacheAccess(b.ctx).then(() => order.push("b in"));
    await new Promise((r) => setTimeout(r, 120));
    order.push("a out");
    await plugin.afterCacheAccess(a.ctx);
    await waiting;
    await plugin.afterCacheAccess(b.ctx);
    expect(order).toEqual(["a out", "b in"]);
    await expect(fs.stat(`${file}.lock`)).rejects.toThrow();
  });

  it("treats an unreadable file as signed out, and a new sign-in replaces it", async () => {
    await new EncryptedFile(file, "the key it was written with").write("old tokens");
    const store = new EncryptedFile(file, KEY);
    const plugin = encryptedCachePlugin(store);
    const { ctx, seen } = context(true);
    await plugin.beforeCacheAccess(ctx);
    expect(seen).toEqual([]);
    await plugin.afterCacheAccess(ctx);
    expect(await store.read()).toBe(SECRET_CACHE);
  });
});

describe("acquireLock", () => {
  it("breaks a lock left behind by a dead process", async () => {
    const lock = path.join(dir, "x.lock");
    await fs.writeFile(lock, "12345");
    const old = new Date(Date.now() - 60_000);
    await fs.utimes(lock, old, old);
    const held = await acquireLock(lock);
    expect(held.locked).toBe(true);
    await held.release();
  });

  it("gives up waiting on a live lock after the timeout, and says so", async () => {
    const lock = path.join(dir, "x.lock");
    const first = await acquireLock(lock);
    const second = await acquireLock(lock, { timeoutMs: 100, delayMs: 20 });
    expect(second.locked).toBe(false);
    await first.release();
  });
});

describe("clearEncryptedTokens", () => {
  const config = (key = KEY) => ({ clientId: "cid", tokenCachePath: file, tokenCacheKey: key });
  const cache = (clients: string[]) =>
    JSON.stringify({
      Account: { a: { home_account_id: "h" } },
      RefreshToken: Object.fromEntries(clients.map((c) => [`rt-${c}`, { client_id: c, home_account_id: "h" }])),
    });

  it("removes the file once nothing is left in it", async () => {
    await new EncryptedFile(file, KEY).write(cache(["cid"]));
    expect(await clearEncryptedTokens(config())).toMatchObject({ status: "cleared", removed: 1, deletedStore: true });
    await expect(fs.stat(file)).rejects.toThrow();
  });

  it("keeps another app's tokens", async () => {
    await new EncryptedFile(file, KEY).write(cache(["cid", "other-app"]));
    expect(await clearEncryptedTokens(config())).toMatchObject({ status: "cleared", removed: 1, deletedStore: false });
    expect(JSON.parse((await new EncryptedFile(file, KEY).read())!).RefreshToken).toHaveProperty("rt-other-app");
  });

  it("leaves a file it cannot decrypt alone, and says how to sign out", async () => {
    expect(await clearEncryptedTokens(config())).toMatchObject({ status: "nothing-stored" });
    await new EncryptedFile(file, KEY).write(cache(["cid"]));
    expect(await clearEncryptedTokens(config("not the key it was written with"))).toMatchObject({
      status: "unavailable",
      reason: expect.stringContaining("delete the file"),
    });
    await expect(fs.stat(file)).resolves.toBeTruthy();
  });
});

describe("credentials with MCP_TOKEN_CACHE_KEY", () => {
  // The point of the key: persistence where the OS-store plugin cannot load.
  it("persist to the encrypted file, with no keyring plugin involved", async () => {
    const read = vi.spyOn(EncryptedFile.prototype, "read");
    const config = {
      authMode: "device-code",
      tenantId: "common",
      clientId: "cid",
      tokenCachePath: file,
      tokenCacheKey: KEY,
      scopes: ["User.Read"],
      enableWrites: false,
      perSurfaceWrites: {},
      disabledTools: new Set<string>(),
      logLevel: "silent",
      listTools: false,
      logout: false,
    } as ServerConfig;
    const record = { authority: "login.microsoftonline.com", homeAccountId: "h.t", clientId: "cid", tenantId: "t", username: "u" };
    const cred = await buildCredential(config, record);
    expect(cred).toBeInstanceOf(DeviceCodeCredential);
    // A silent request looks the account up in the cache: through our plugin.
    await expect(cred.getToken("https://graph.microsoft.com/.default")).rejects.toMatchObject({ name: "AuthenticationRequiredError" });
    expect(read).toHaveBeenCalled();
  });
});
