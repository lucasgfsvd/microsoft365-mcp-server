import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import type { TokenCredential } from "@azure/identity";
import type { z } from "zod";
import type { AuthSession } from "./auth/session.js";

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
  /** Largest single MCP message accepted on stdin. Bounds how big a base64 upload can be. */
  maxMessageBytes: number;
}

export interface ToolContext {
  graph: GraphClient;
  credential: TokenCredential;
  config: ServerConfig;
  auth: AuthSession;
}

/**
 * Any zod schema. The `any` type arguments are deliberate — this is the
 * "accepts any schema" constraint for ToolDefinition, and narrowing them makes
 * z.infer resolve to `unknown` in every tool handler. (zod 4 replaced the old
 * z.ZodTypeAny, whose bare z.ZodType successor defaults to `unknown`.)
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type ZodObj = z.ZodType<any, any>;

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
  | "powerpoint"
  | "auth"
  | "graph";

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
  "auth",
  "graph",
];
