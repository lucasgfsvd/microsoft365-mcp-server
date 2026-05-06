import { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";
import type { TokenCredential } from "@azure/identity";

/**
 * Build a Microsoft Graph client wired to the given Azure credential.
 *
 * `Client.initWithMiddleware({ authProvider })` constructs the SDK's default
 * middleware chain, which already includes a `RetryHandler` configured with
 * `RetryHandlerOptions` defaults: 3 retries, 3-second base delay, exponential
 * backoff, honors the `Retry-After` header on 429/503/504. Tools therefore
 * don't need to implement their own retry logic — Graph throttling is handled
 * transparently up to that ceiling. Beyond that, errors propagate as
 * `GraphError` and are surfaced via `normalizeGraphError`.
 */
export function buildGraphClient(credential: TokenCredential, scopes: string[]): GraphClient {
  const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes });
  return GraphClient.initWithMiddleware({
    authProvider,
    defaultVersion: "v1.0",
  });
}
