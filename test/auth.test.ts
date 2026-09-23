import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import type { AccessToken, TokenCredential } from "@azure/identity";
import type { ServerConfig } from "../src/types.js";
import { AuthSession } from "../src/auth/session.js";
import { authRecordPath } from "../src/auth/tokenCache.js";
import { runLogout } from "../src/cli.js";

function makeConfig(over: Partial<ServerConfig> = {}): ServerConfig {
  return {
    authMode: "device-code",
    tenantId: "common",
    clientId: "cid",
    tokenCachePath: "/tmp/tokencache.json",
    scopes: ["User.Read"],
    enableWrites: false,
    perSurfaceWrites: {},
    disabledTools: new Set<string>(),
    logLevel: "silent",
    listTools: false,
    logout: false,
    ...over,
  } as ServerConfig;
}

/** Credential whose behaviour can be flipped mid-test. */
function flakyCredential() {
  const state = { ok: true };
  const cred: TokenCredential = {
    getToken: async () => {
      if (!state.ok) {
        const e = new Error("Automatic authentication has been disabled.");
        e.name = "AuthenticationRequiredError";
        throw e;
      }
      return { token: "t", expiresOnTimestamp: Date.now() + 3_600_000 } as AccessToken;
    },
  };
  return { cred, state };
}

describe("AuthSession.status", () => {
  // Regression: status() used to delegate to probe(), which short-circuits on
  // the cached flag. Once signed in it reported signedIn: true forever, even
  // while every Graph call was failing — sending you after the wrong bug.
  it("re-checks rather than trusting a stale signed-in flag", async () => {
    const { cred, state } = flakyCredential();
    const auth = new AuthSession(cred, makeConfig());

    expect(await auth.probe()).toBe(true);
    expect((await auth.status()).signedIn).toBe(true);

    state.ok = false; // the credential can no longer produce a token

    const s = await auth.status();
    expect(s.signedIn).toBe(false);
    expect(s.state).toBe("sign-in-required");
  });

  it("recovers once the credential works again", async () => {
    const { cred, state } = flakyCredential();
    const auth = new AuthSession(cred, makeConfig());
    state.ok = false;
    expect((await auth.status()).signedIn).toBe(false);
    state.ok = true;
    expect((await auth.status()).signedIn).toBe(true);
  });
});

describe("runLogout", () => {
  let dir: string;
  beforeEach(async () => {
    dir = await fs.mkdtemp(path.join(os.tmpdir(), "m365-logout-"));
  });
  afterEach(async () => {
    await fs.rm(dir, { recursive: true, force: true });
  });

  // Regression: logout removed the token cache but left authrecord.json, so a
  // restart still believed it knew the account whose token had just been deleted.
  it("removes the authentication record as well as the token cache", async () => {
    const cachePath = path.join(dir, "tokencache.json");
    await fs.writeFile(cachePath, "{}");
    await fs.writeFile(authRecordPath(cachePath), "{}");

    await runLogout(makeConfig({ tokenCachePath: cachePath }));

    await expect(fs.access(cachePath)).rejects.toThrow();
    await expect(fs.access(authRecordPath(cachePath))).rejects.toThrow();
  });

  it("is idempotent when nothing is there", async () => {
    const cachePath = path.join(dir, "tokencache.json");
    await expect(runLogout(makeConfig({ tokenCachePath: cachePath }))).resolves.toBeUndefined();
  });

  it("still clears the record when only the record remains", async () => {
    const cachePath = path.join(dir, "tokencache.json");
    await fs.writeFile(authRecordPath(cachePath), "{}");
    await runLogout(makeConfig({ tokenCachePath: cachePath }));
    await expect(fs.access(authRecordPath(cachePath))).rejects.toThrow();
  });
});
