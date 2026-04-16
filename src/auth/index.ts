import { EventEmitter } from "node:events";
import {
  ClientSecretCredential,
  DeviceCodeCredential,
  InteractiveBrowserCredential,
  useIdentityPlugin,
  type TokenCredential,
} from "@azure/identity";
import { cachePersistencePlugin } from "@azure/identity-cache-persistence";
import { logger } from "../util/logger.js";
import type { ServerConfig } from "../types.js";
import { defaultCachePath } from "./tokenCache.js";

let cachePluginRegistered = false;
function ensureCachePlugin(): void {
  if (cachePluginRegistered) return;
  useIdentityPlugin(cachePersistencePlugin);
  cachePluginRegistered = true;
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
export function buildCredential(config: ServerConfig): TokenCredential {
  const { authMode, tenantId, clientId, clientSecret, redirectUri } = config;
  const tokenCachePersistenceOptions = {
    enabled: true,
    name: "microsoft365-mcp",
    unsafeAllowUnencryptedStorage: true,
  } as const;

  switch (authMode) {
    case "device-code":
      ensureCachePlugin();
      logger.info({ tenantId, clientId }, "auth: using device code flow");
      return new DeviceCodeCredential({
        tenantId,
        clientId,
        tokenCachePersistenceOptions,
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

    case "client-credentials":
      if (!clientSecret) {
        throw new Error(
          "MCP_CLIENT_SECRET is required for client-credentials auth mode. " +
            "Set it in the environment or switch to --auth device-code.",
        );
      }
      logger.info({ tenantId, clientId }, "auth: using client credentials flow");
      return new ClientSecretCredential(tenantId, clientId, clientSecret);

    case "interactive":
      ensureCachePlugin();
      logger.info({ tenantId, clientId }, "auth: using interactive browser flow");
      return new InteractiveBrowserCredential({
        tenantId,
        clientId,
        redirectUri: redirectUri ?? "http://localhost:3000",
        tokenCachePersistenceOptions,
      });
  }
}

export { defaultCachePath };
