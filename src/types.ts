import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import type { TokenCredential } from "@azure/identity";
import type { z } from "zod";

export type AuthMode = "device-code" | "client-credentials" | "interactive";

export interface ServerConfig {
  authMode: AuthMode;
  tenantId: string;
  clientId: string;
  clientSecret?: string;
  redirectUri?: string;
  tokenCachePath: string;
  scopes: string[];
  enableWrites: boolean;
  perSurfaceWrites: Record<string, boolean>;
  disabledTools: Set<string>;
  logLevel: string;
  /** When true, print the tool catalogue to stdout and exit before starting the server. */
  listTools: boolean;
  /** When true, delete the cached OAuth tokens and exit before starting the server. */
  logout: boolean;
}

export interface ToolContext {
  graph: GraphClient;
  credential: TokenCredential;
  config: ServerConfig;
}

export type ZodObj = z.ZodTypeAny;

export interface ToolDefinition<I extends ZodObj = ZodObj> {
  name: string;
  description: string;
  surface: Surface;
  inputSchema: I;
  /** true for tools that create, update, delete or send. */
  mutating?: boolean;
  /** Graph scopes required to execute this tool. */
  requiredScopes?: string[];
  handler: (input: z.infer<I>, ctx: ToolContext) => Promise<unknown>;
}

export type Surface =
  | "mail"
  | "calendar"
  | "contacts"
  | "files"
  | "teams"
  | "tasks"
  | "onenote"
  | "excel"
  | "word"
  | "powerpoint";

export const SURFACES: Surface[] = [
  "mail",
  "calendar",
  "contacts",
  "files",
  "teams",
  "tasks",
  "onenote",
  "excel",
  "word",
  "powerpoint",
];
