import "isomorphic-fetch";
import { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";
import type { TokenCredential } from "@azure/identity";

export function buildGraphClient(credential: TokenCredential, scopes: string[]): GraphClient {
  const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes });
  return GraphClient.initWithMiddleware({
    authProvider,
    defaultVersion: "v1.0",
  });
}
