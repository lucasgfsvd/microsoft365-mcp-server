import { describe, it, expect, beforeEach } from "vitest";
import { DeviceCodeCredential } from "@azure/identity";
import type { ServerConfig } from "../src/types.js";
import {
  buildCredential,
  ensureCachePlugin,
  resetCachePluginForTests,
} from "../src/auth/index.js";

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
