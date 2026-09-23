import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import type { ServerConfig } from "../src/types.js";
import { authRecordPath } from "../src/auth/tokenCache.js";
import { clearStoredTokens, pruneClient, type ClearOutcome } from "../src/auth/tokenStore.js";
import { runLogout } from "../src/cli.js";

const OURS = "our-client";
const THEIRS = "other-app";

/** An MSAL serialized cache: one account, with tokens from us and optionally another app. */
function cacheWith(opts: { otherApp?: boolean } = {}) {
  const home = "uid.tid";
  const cred = (client: string, kind: string) => ({
    [`${home}-login.microsoftonline.com-${kind.toLowerCase()}-${client}`]: {
      home_account_id: home,
      client_id: client,
      credential_type: kind,
    },
  });
  return JSON.stringify({
    Account: { [`${home}-login.microsoftonline.com-tid`]: { home_account_id: home } },
    AccessToken: { ...cred(OURS, "AccessToken"), ...(opts.otherApp ? cred(THEIRS, "AccessToken") : {}) },
    RefreshToken: { ...cred(OURS, "RefreshToken"), ...(opts.otherApp ? cred(THEIRS, "RefreshToken") : {}) },
    IdToken: cred(OURS, "IdToken"),
    AppMetadata: { [`appmetadata-login.microsoftonline.com-${OURS}`]: { client_id: OURS } },
  });
}

describe("pruneClient", () => {
  it("removes every credential of ours, and the store is then empty", () => {
    const r = pruneClient(cacheWith(), OURS);
    expect(r.removed).toBe(3);
    expect(r.empty).toBe(true);
  });

  // The keychain item on macOS/Linux is shared by every app using the plugin.
  it("keeps another app's tokens and the account they belong to", () => {
    const r = pruneClient(cacheWith({ otherApp: true }), OURS);
    const left = JSON.parse(r.json);
    expect(r.removed).toBe(3);
    expect(r.empty).toBe(false);
    expect(Object.values(left.RefreshToken)).toEqual([expect.objectContaining({ client_id: THEIRS })]);
    expect(Object.keys(left.Account)).toHaveLength(1);
    expect(left.AppMetadata).toEqual({});
  });

  it("removes nothing when none of it is ours", () => {
    expect(pruneClient(cacheWith(), "someone-else").removed).toBe(0);
  });
});

describe("clearStoredTokens", () => {
  it("reports unavailable when the persistence layer cannot load (no libsecret)", async () => {
    const r = await clearStoredTokens({ clientId: OURS, tokenCachePath: "/x/tokencache.json" }, () =>
      Promise.reject(new Error("libsecret-1.so.0: cannot open shared object file")),
    );
    expect(r).toEqual({ status: "unavailable", reason: expect.stringContaining("libsecret") });
  });
});

describe("runLogout", () => {
  let dir: string;
  let cachePath: string;
  const config = () => ({ clientId: OURS, tokenCachePath: cachePath }) as ServerConfig;
  const cleared = (outcome: ClearOutcome) => {
    const calls: unknown[] = [];
    const fn = async (c: unknown) => {
      calls.push(c);
      return outcome;
    };
    return { fn: fn as typeof clearStoredTokens, calls };
  };

  beforeEach(async () => {
    dir = await fs.mkdtemp(path.join(os.tmpdir(), "m365-logout-"));
    cachePath = path.join(dir, "tokencache.json");
  });
  afterEach(async () => {
    await fs.rm(dir, { recursive: true, force: true });
  });

  // Regression: logout removed only authrecord.json and a tokencache.json that
  // was never written, leaving a usable refresh token in the real store.
  it("clears the token store as well as the record", async () => {
    await fs.writeFile(authRecordPath(cachePath), "{}");
    const clear = cleared({ status: "cleared", removed: 3, deletedStore: true, location: "store" });

    await runLogout(config(), clear.fn);

    expect(clear.calls).toHaveLength(1);
    await expect(fs.access(authRecordPath(cachePath))).rejects.toThrow();
  });

  it("is idempotent when nothing is there", async () => {
    const clear = cleared({ status: "nothing-stored", location: "store" });
    await expect(runLogout(config(), clear.fn)).resolves.toBeUndefined();
  });

  it("still removes the record when there is no persistent store", async () => {
    await fs.writeFile(authRecordPath(cachePath), "{}");
    await runLogout(config(), cleared({ status: "unavailable", reason: "no libsecret" }).fn);
    await expect(fs.access(authRecordPath(cachePath))).rejects.toThrow();
  });
});
