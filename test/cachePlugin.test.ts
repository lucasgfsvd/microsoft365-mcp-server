import { describe, it, expect, beforeEach } from "vitest";
import os from "node:os";
import path from "node:path";
import { DeviceCodeCredential } from "@azure/identity";
import type { ServerConfig } from "../src/types.js";
import {
  buildCredential,
  ensureCachePlugin,
  resetCachePluginForTests,
} from "../src/auth/index.js";
import { defaultCachePath, persistentCacheName } from "../src/auth/tokenCache.js";

const config = {
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
} as ServerConfig;

// Regression: the cache plugin was imported statically, and it loads keytar,
// which dlopens libsecret. On a host without libsecret the auth module could
// not even be imported, despite keytar being declared optional.
describe("persistent cache plugin", () => {
  beforeEach(() => resetCachePluginForTests());

  it("degrades to an in-memory cache when the plugin cannot load", async () => {
    const dlopenFailure = Object.assign(new Error("libsecret-1.so.0: cannot open shared object"), {
      code: "ERR_DLOPEN_FAILED",
    });
    expect(await ensureCachePlugin(() => Promise.reject(dlopenFailure))).toBe(false);

    // Requesting persistence without a registered provider makes @azure/identity
    // throw, so the credential must be built without it.
    await expect(buildCredential(config)).resolves.toBeInstanceOf(DeviceCodeCredential);
  });

  it("loads the plugin once, however many credentials are built", async () => {
    let loads = 0;
    const load = () => {
      loads++;
      return Promise.reject(new Error("unavailable"));
    };
    await ensureCachePlugin(load);
    await ensureCachePlugin(load);
    expect(loads).toBe(1);
  });

  it("does not touch the plugin for client credentials", async () => {
    let loads = 0;
    await buildCredential({ ...config, authMode: "client-credentials", clientSecret: "s" });
    await ensureCachePlugin(() => {
      loads++;
      return Promise.reject(new Error("unavailable"));
    });
    expect(loads).toBe(1);
  });
});

// Regression: the plugin's cache name was fixed, so MCP_TOKEN_CACHE_PATH moved
// only the auth record while every server shared one token store.
describe("persistentCacheName", () => {
  it("keeps the original name for the default path, so existing sign-ins survive", () => {
    expect(persistentCacheName(defaultCachePath())).toBe("microsoft365-mcp");
  });

  it("gives other paths their own stable cache", () => {
    const a = persistentCacheName(path.join(os.tmpdir(), "a", "tokencache.json"));
    const b = persistentCacheName(path.join(os.tmpdir(), "b", "tokencache.json"));
    expect(a).toMatch(/^microsoft365-mcp-[0-9a-f]{12}$/);
    expect(a).not.toBe(b);
    expect(persistentCacheName(path.join(os.tmpdir(), "a", "tokencache.json"))).toBe(a);
  });
});
