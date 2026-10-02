import { describe, it, expect } from "vitest";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { ClientSecretCredential, type AccessToken, type AuthenticationRecord, type TokenCredential } from "@azure/identity";
import type { ServerConfig } from "../src/types.js";
import { buildServerCredential, ReloadingCredential } from "../src/auth/reloadingCredential.js";
import { writeAuthRecord } from "../src/auth/tokenCache.js";

const authRequired = () => Object.assign(new Error("Automatic authentication has been disabled"), { name: "AuthenticationRequiredError" });
const token = (t: string): AccessToken => ({ token: t, expiresOnTimestamp: Date.now() + 3_600_000 });
const record = (homeAccountId: string, clientId = "cid"): AuthenticationRecord => ({
  authority: "login.microsoftonline.com",
  homeAccountId,
  clientId,
  tenantId: "t",
  username: "u@example.com",
});

/** A credential that only has an account when built with a record. */
const credentialFor = (r?: AuthenticationRecord): TokenCredential & { authenticate: () => Promise<AuthenticationRecord> } => ({
  getToken: async () => {
    if (!r) throw authRequired();
    return token(`token-for-${r.homeAccountId}`);
  },
  authenticate: async () => record("signed-in-here"),
});

function setup(onDisk: AuthenticationRecord | undefined, start?: AuthenticationRecord) {
  const builds: Array<AuthenticationRecord | undefined> = [];
  let disk = onDisk;
  const cred = new ReloadingCredential(
    credentialFor(start),
    start,
    async (r) => (builds.push(r), credentialFor(r)),
    async () => disk,
    "cid",
  );
  return { cred, builds, setDisk: (r?: AuthenticationRecord) => (disk = r) };
}

describe("ReloadingCredential", () => {
  it("picks up a record another process wrote, without a restart", async () => {
    const { cred, builds } = setup(record("elsewhere"));
    expect((await cred.getToken("scope"))?.token).toBe("token-for-elsewhere");
    expect(builds).toEqual([record("elsewhere")]);
    // Rebuilt once; later calls use the new credential directly.
    await cred.getToken("scope");
    expect(builds).toHaveLength(1);
  });

  it("still asks for sign-in when there is nothing new on disk", async () => {
    for (const disk of [undefined, record("x", "another-app")]) {
      const { cred, builds } = setup(disk);
      await expect(cred.getToken("scope")).rejects.toThrow(/Automatic authentication/);
      expect(builds).toEqual([]);
    }
  });

  it("does not rebuild for the record it already has", async () => {
    const r = record("same");
    const builds: unknown[] = [];
    const failing: TokenCredential = { getToken: async () => Promise.reject(authRequired()) };
    const cred = new ReloadingCredential(failing, r, async (x) => (builds.push(x), failing), async () => ({ ...r }), "cid");
    await expect(cred.getToken("scope")).rejects.toThrow(/Automatic authentication/);
    expect(builds).toEqual([]);
  });

  it("passes other failures straight through, without reading the disk", async () => {
    let reads = 0;
    const broken: TokenCredential = { getToken: async () => Promise.reject(new Error("ENOTFOUND login.microsoftonline.com")) };
    const cred = new ReloadingCredential(broken, undefined, async () => broken, async () => (reads++, record("x")), "cid");
    await expect(cred.getToken("scope")).rejects.toThrow(/ENOTFOUND/);
    expect(reads).toBe(0);
  });

  it("reloads once for concurrent callers", async () => {
    const { cred, builds } = setup(record("elsewhere"));
    const tokens = await Promise.all([cred.getToken("a"), cred.getToken("b"), cred.getToken("c")]);
    expect(tokens.map((t) => t?.token)).toEqual(Array(3).fill("token-for-elsewhere"));
    expect(builds).toHaveLength(1);
  });

  it("signs in through the current credential and remembers the record", async () => {
    const { cred, builds } = setup(record("signed-in-here"));
    expect(await cred.authenticate("scope")).toEqual(record("signed-in-here"));
    // The record on disk is the one this process just made: nothing to pick up.
    const inner = (cred as unknown as { inner: TokenCredential }).inner;
    inner.getToken = async () => Promise.reject(authRequired());
    await expect(cred.getToken("scope")).rejects.toThrow();
    expect(builds).toEqual([]);
  });
});

describe("buildServerCredential", () => {
  const base = {
    tenantId: "common",
    clientId: "cid",
    scopes: ["User.Read"],
    enableWrites: false,
    perSurfaceWrites: {},
    disabledTools: new Set<string>(),
    logLevel: "silent",
    listTools: false,
    logout: false,
  };

  it("wraps user sign-in, starting from the record on disk", async () => {
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), "m365-reload-"));
    try {
      const tokenCachePath = path.join(dir, "tokencache.json");
      await writeAuthRecord(tokenCachePath, record("from-disk"));
      // A cache key keeps the OS keyring plugin out of it: on a CI runner without libsecret,
      // loading it raised an unhandled error under Node 20.
      const cred = await buildServerCredential({ ...base, authMode: "device-code", tokenCachePath, tokenCacheKey: "a key for the test only" } as ServerConfig);
      expect(cred).toBeInstanceOf(ReloadingCredential);
      expect((cred as unknown as { record: AuthenticationRecord }).record).toEqual(record("from-disk"));
    } finally {
      await fs.rm(dir, { recursive: true, force: true });
    }
  });

  it("leaves client credentials alone: there is no account to pick up", async () => {
    const cred = await buildServerCredential({ ...base, authMode: "client-credentials", clientSecret: "s", tokenCachePath: path.join(os.tmpdir(), "none", "tc.json") } as ServerConfig);
    expect(cred).toBeInstanceOf(ClientSecretCredential);
  });
});
