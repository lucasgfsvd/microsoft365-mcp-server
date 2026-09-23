import { EventEmitter } from "node:events";
import type { AuthenticationRecord } from "@azure/identity";
import {
  ClientSecretCredential,
  DeviceCodeCredential,
  InteractiveBrowserCredential,
  useIdentityPlugin,
  type TokenCredential,
} from "@azure/identity";
import { logger } from "../util/logger.js";
import type { ServerConfig } from "../types.js";
import { defaultCachePath, persistentCacheName } from "./tokenCache.js";

let cachePlugin: Promise<boolean> | undefined;

/**
 * Register the persistent token cache, reporting whether it is available.
 *
 * Loaded lazily because the plugin pulls in `keytar` at module load, and keytar
 * dlopens libsecret on Linux. Where libsecret is missing (headless hosts, the
 * distroless image) a static import would take the whole process down before
 * any fallback could run. Instead we degrade to an in-memory cache: sign-in
 * still works, it just does not survive a restart.
 */
export function ensureCachePlugin(
  load: () => Promise<{ cachePersistencePlugin: Parameters<typeof useIdentityPlugin>[0] }> = () =>
    import("@azure/identity-cache-persistence"),
): Promise<boolean> {
  cachePlugin ??= load().then(
    ({ cachePersistencePlugin }) => {
      useIdentityPlugin(cachePersistencePlugin);
      return true;
    },
    (err: unknown) => {
      logger.warn(
        { err: err instanceof Error ? err.message : String(err) },
        "auth: persistent token cache unavailable; tokens will be held in memory only " +
          "and sign-in will not survive a restart",
      );
      return false;
    },
  );
  return cachePlugin;
}

/** Test seam: forget the memoised plugin load. */
export function resetCachePluginForTests(): void {
  cachePlugin = undefined;
}

export interface DeviceCodePrompt {
  verificationUri: string;
  userCode: string;
  message?: string;
}

let _lastDeviceCodePrompt: DeviceCodePrompt | undefined;
export function getLastDeviceCodePrompt(): DeviceCodePrompt | undefined {
  return _lastDeviceCodePrompt;
}

/** Drop a stale code so a new sign-in attempt cannot report the previous one. */
export function clearLastDeviceCodePrompt(): void {
  _lastDeviceCodePrompt = undefined;
}

export const deviceCodeEmitter = new EventEmitter();

/**
 * Build a TokenCredential suitable for Microsoft Graph based on the configured auth mode.
 *
 * - device-code: prints a user code to stderr; user visits microsoft.com/devicelogin.
 * - client-credentials: unattended; requires tenant + client id + secret and admin-consented app perms.
 * - interactive: opens a local browser for auth code + PKCE.
 */
export async function buildCredential(
  config: ServerConfig,
  authenticationRecord?: AuthenticationRecord,
): Promise<TokenCredential> {
  const { authMode, tenantId, clientId, clientSecret, redirectUri } = config;
  // Only ask for persistence when the plugin actually loaded: requesting it with
  // no provider registered makes @azure/identity throw.
  const persist = async () =>
    (await ensureCachePlugin())
      ? ({
          enabled: true,
          name: persistentCacheName(config.tokenCachePath),
          unsafeAllowUnencryptedStorage: true,
        } as const)
      : undefined;

  switch (authMode) {
    case "device-code": {
      const tokenCachePersistenceOptions = await persist();
      logger.info({ tenantId, clientId }, "auth: using device code flow");
      return new DeviceCodeCredential({
        tenantId,
        clientId,
        tokenCachePersistenceOptions,
        authenticationRecord,
        // Never prompt from getToken(). Interactive sign-in happens only through
        // AuthSession.signIn(), so a client launch stays silent.
        disableAutomaticAuthentication: true,
        userPromptCallback: (info) => {
          const prompt: DeviceCodePrompt = {
            verificationUri: info.verificationUri,
            userCode: info.userCode,
            message: info.message,
          };
          _lastDeviceCodePrompt = prompt;
          deviceCodeEmitter.emit("prompt", prompt);
          // Also write to stderr so CLI users / log-tailers see it.
          // stdout is reserved for MCP protocol.
          process.stderr.write(
            `\n[microsoft365-mcp] To sign in, open ${info.verificationUri} and enter code ${info.userCode}\n\n`,
          );
        },
      });
    }

    case "client-credentials":
      if (!clientSecret) {
        throw new Error(
          "MCP_CLIENT_SECRET is required for client-credentials auth mode. " +
            "Set it in the environment or switch to --auth device-code.",
        );
      }
      logger.info({ tenantId, clientId }, "auth: using client credentials flow");
      return new ClientSecretCredential(tenantId, clientId, clientSecret);

    case "interactive": {
      const tokenCachePersistenceOptions = await persist();
      logger.info({ tenantId, clientId }, "auth: using interactive browser flow");
      return new InteractiveBrowserCredential({
        tenantId,
        clientId,
        redirectUri: redirectUri ?? "http://localhost:3000",
        tokenCachePersistenceOptions,
        authenticationRecord,
        disableAutomaticAuthentication: true,
      });
    }
  }
}

export { defaultCachePath };
