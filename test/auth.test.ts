import { describe, it, expect } from "vitest";
import type { AccessToken, TokenCredential } from "@azure/identity";
import type { ServerConfig } from "../src/types.js";
import { AuthSession } from "../src/auth/session.js";

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
