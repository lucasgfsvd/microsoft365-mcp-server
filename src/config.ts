import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { Command } from "commander";
import type { AuthMode, ServerConfig } from "./types.js";
import { SURFACES } from "./types.js";
import { defaultCachePath } from "./auth/tokenCache.js";

function readPackageVersion(): string {
  try {
    const here = dirname(fileURLToPath(import.meta.url));
    const pkg = JSON.parse(readFileSync(resolve(here, "../package.json"), "utf8")) as { version: string };
    return pkg.version;
  } catch {
    return "0.0.0";
  }
}

// Default to read-only scopes — least privilege. Write scopes are added below
// when MCP_ENABLE_WRITES (or any per-surface write flag) is set, matching the
// writeGuard's default-deny posture. Set MCP_SCOPES explicitly to fully override.
const READ_SCOPES = [
  "offline_access",
  "User.Read",
  "Mail.Read",
  "Calendars.Read",
  "Calendars.Read.Shared",
  "Contacts.Read",
  "People.Read",
  "Files.Read.All",
  "Sites.Read.All",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "Chat.Read",
  "Tasks.Read",
  "Notes.Read",
];

const WRITE_SCOPES = [
  "Mail.Send",
  "Mail.ReadWrite",
  "Calendars.ReadWrite",
  "Contacts.ReadWrite",
  "Files.ReadWrite.All",
  "Sites.ReadWrite.All",
  "ChannelMessage.Send",
  "ChatMessage.Send",
  "Tasks.ReadWrite",
  "Notes.ReadWrite",
];

/** Microsoft's "Microsoft Graph Command Line Tools" public client.
 *  Convenient for trying the server out; override MCP_CLIENT_ID for production
 *  so you get clean audit logs and can scope your own permissions. */
export const DEFAULT_PUBLIC_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";

function envBool(name: string): boolean | undefined {
  const v = process.env[name];
  if (v === undefined) return undefined;
  return /^(1|true|yes|on)$/i.test(v);
}

export function loadConfig(argv: string[]): ServerConfig {
  const program = new Command()
    .name("microsoft365-mcp-server")
    .description("Model Context Protocol server for Microsoft 365")
    .version(readPackageVersion())
    .option("--auth <mode>", "Auth mode: device-code | client-credentials | interactive")
    .option("--tenant <id>", "Azure AD tenant id (or 'common', 'organizations', 'consumers')")
    .option("--client-id <id>", "Azure AD application (client) id")
    .option("--client-secret <secret>", "Client secret (client-credentials only)")
    .option("--redirect-uri <url>", "Redirect URI (interactive only)")
    .option("--scopes <csv>", "Comma-separated Graph scopes (overrides default)")
    .option("--enable-writes", "Enable all mutating tools")
    .option("--disabled-tools <csv>", "Comma-separated tool names to hide")
    .option("--token-cache <path>", "Token cache file path")
    .option("--list-tools", "Print the tool catalogue and exit (no server, no auth)")
    .option("--logout", "Delete the cached OAuth token and exit (forces re-auth on next run)")
    .allowExcessArguments(true)
    .allowUnknownOption(true)
    .parse(argv);

  const opts = program.opts();

  const authMode = (opts.auth ?? process.env.MCP_AUTH_MODE ?? "device-code") as AuthMode;
  if (!["device-code", "client-credentials", "interactive"].includes(authMode)) {
    throw new Error(`Invalid auth mode: ${authMode}`);
  }

  const tenantId = opts.tenant ?? process.env.MCP_TENANT_ID ?? "common";
  const clientId =
    opts.clientId ?? process.env.MCP_CLIENT_ID ?? DEFAULT_PUBLIC_CLIENT_ID;
  const clientSecret = opts.clientSecret ?? process.env.MCP_CLIENT_SECRET;
  const redirectUri = opts.redirectUri ?? process.env.MCP_REDIRECT_URI;

  const enableWrites = Boolean(opts.enableWrites) || envBool("MCP_ENABLE_WRITES") === true;

  const perSurfaceWrites: Record<string, boolean> = {};
  for (const s of SURFACES) {
    const envName = `MCP_ENABLE_${s.toUpperCase()}_WRITE`;
    const v = envBool(envName);
    if (v !== undefined) perSurfaceWrites[s] = v;
  }

  const wantsWrites = enableWrites || Object.values(perSurfaceWrites).some(Boolean);
  const defaultScopes = wantsWrites ? [...READ_SCOPES, ...WRITE_SCOPES] : READ_SCOPES;
  const scopes = (opts.scopes ?? process.env.MCP_SCOPES ?? defaultScopes.join(","))
    .split(",")
    .map((s: string) => s.trim())
    .filter(Boolean);

  // For client-credentials flow, Graph requires the single ".default" scope.
  const effectiveScopes = authMode === "client-credentials" ? ["https://graph.microsoft.com/.default"] : scopes;

  const disabledTools = new Set<string>(
    (opts.disabledTools ?? process.env.MCP_DISABLED_TOOLS ?? "")
      .split(",")
      .map((s: string) => s.trim())
      .filter(Boolean),
  );

  return {
    authMode,
    tenantId,
    clientId,
    clientSecret,
    redirectUri,
    tokenCachePath: opts.tokenCache ?? process.env.MCP_TOKEN_CACHE_PATH ?? defaultCachePath(),
    scopes: effectiveScopes,
    enableWrites,
    perSurfaceWrites,
    disabledTools,
    logLevel: process.env.MCP_LOG_LEVEL ?? "info",
    listTools: Boolean(opts.listTools),
    logout: Boolean(opts.logout),
  };
}
