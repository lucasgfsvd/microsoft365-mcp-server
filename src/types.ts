import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import type { TokenCredential } from "@azure/identity";
import type { z } from "zod";
import type { AuthSession } from "./auth/session.js";
import type { RetryPolicy } from "./graph/retry.js";
import type { ProgressReporter } from "./util/progress.js";

export type AuthMode = "device-code" | "client-credentials" | "interactive";

export interface ServerConfig {
  authMode: AuthMode;
  tenantId: string;
  clientId: string;
  clientSecret?: string;
  /** PEM file holding the certificate and its private key (client-credentials). */
  clientCertificatePath?: string;
  clientCertificatePassword?: string;
  /**
   * The user app-only mode works for: whose mail, calendar and files `/me`
   * stands for. Required with client-credentials, which has no signed-in user.
   */
  user?: string;
  /** Offer a `mailbox` argument on mail and calendar tools (shared mailboxes, delegates). */
  sharedMailboxes: boolean;
  /** Request the scopes only an admin can grant, and offer the tools that need them. */
  adminScopes: boolean;
  redirectUri?: string;
  tokenCachePath: string;
  /**
   * Secret for an encrypted token cache file at tokenCachePath, used instead of
   * the OS store. From MCP_TOKEN_CACHE_KEY only: never a flag, never logged.
   */
  tokenCacheKey?: string;
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
  /** Where files_download may write with saveToDisk. Unset disables writing to disk. */
  downloadDir?: string;
  /** The only folder files_upload may read a localPath from. Unset disables local uploads. */
  uploadDir?: string;
}

export interface ToolContext {
  graph: GraphClient;
  credential: TokenCredential;
  config: ServerConfig;
  auth: AuthSession;
  /** Set when the client asked for progress; long transfers report through it. */
  progress?: ProgressReporter;
  /** Aborted when the client cancels the call. */
  signal?: AbortSignal;
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
  /** Retry tuning; mutating POSTs are only ever retried on 429 regardless. */
  retry?: RetryPolicy;
  /**
   * Needs a permission only an admin can grant (rooms, others'
   * full profiles). Hidden unless MCP_ENABLE_ADMIN_SCOPES, since asking for such
   * a scope at sign-in fails outright where no admin has approved it.
   */
  needsAdminConsent?: boolean;
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
