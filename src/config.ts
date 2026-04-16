import { Command } from "commander";
import type { AuthMode, ServerConfig, Surface } from "./types.js";
import { SURFACES } from "./types.js";
import { defaultCachePath } from "./auth/tokenCache.js";

const DEFAULT_SCOPES = [
  "offline_access",
  "User.Read",
  "Mail.Read",
  "Mail.Send",
  "Mail.ReadWrite",
  "Calendars.ReadWrite",
  "Calendars.Read.Shared",
  "Contacts.ReadWrite",
  "People.Read",
  "Files.ReadWrite.All",
  "Sites.ReadWrite.All",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "ChannelMessage.Send",
  "Chat.Read",
  "ChatMessage.Send",
  "Tasks.ReadWrite",
  "Notes.ReadWrite",
];

function envBool(name: string): boolean | undefined {
  const v = process.env[name];
  if (v === undefined) return undefined;
  return /^(1|true|yes|on)$/i.test(v);
}

export function loadConfig(argv: string[]): ServerConfig {
  const program = new Command()
    .name("microsoft365-mcp-server")
    .description("Model Context Protocol server for Microsoft 365")
    .option("--auth <mode>", "Auth mode: device-code | client-credentials | interactive")
    .option("--tenant <id>", "Azure AD tenant id (or 'common', 'organizations', 'consumers')")
    .option("--client-id <id>", "Azure AD application (client) id")
    .option("--client-secret <secret>", "Client secret (client-credentials only)")
    .option("--redirect-uri <url>", "Redirect URI (interactive only)")
    .option("--scopes <csv>", "Comma-separated Graph scopes (overrides default)")
    .option("--enable-writes", "Enable all mutating tools")
    .option("--disabled-tools <csv>", "Comma-separated tool names to hide")
    .option("--token-cache <path>", "Token cache file path")
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
    opts.clientId ?? process.env.MCP_CLIENT_ID ?? "14d82eec-204b-4c2f-b7e8-296a70dab67e"; // public Graph PowerShell fallback — override in production
  const clientSecret = opts.clientSecret ?? process.env.MCP_CLIENT_SECRET;
  const redirectUri = opts.redirectUri ?? process.env.MCP_REDIRECT_URI;

  const scopes = (opts.scopes ?? process.env.MCP_SCOPES ?? DEFAULT_SCOPES.join(","))
    .split(",")
    .map((s: string) => s.trim())
    .filter(Boolean);

  // For client-credentials flow, Graph requires the single ".default" scope.
  const effectiveScopes = authMode === "client-credentials" ? ["https://graph.microsoft.com/.default"] : scopes;

  const enableWrites = Boolean(opts.enableWrites) || envBool("MCP_ENABLE_WRITES") === true;

  const perSurfaceWrites: Record<string, boolean> = {};
  for (const s of SURFACES) {
    const envName = `MCP_ENABLE_${s.toUpperCase()}_WRITE`;
    const v = envBool(envName);
    if (v !== undefined) perSurfaceWrites[s] = v;
  }

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
  };
}
