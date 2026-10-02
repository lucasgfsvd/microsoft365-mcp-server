import path from "node:path";
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { Command } from "commander";
import type { AuthMode, ServerConfig } from "./types.js";
import { SURFACES } from "./types.js";
import { defaultCachePath } from "./auth/tokenCache.js";

/** The version in package.json, which a release tag must match. */
export function readPackageVersion(): string {
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
  "MailboxSettings.Read",
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
  "MailboxSettings.ReadWrite",
  "Calendars.ReadWrite",
  "Contacts.ReadWrite",
  "Files.ReadWrite.All",
  "Sites.ReadWrite.All",
  "ChannelMessage.Send",
  "ChatMessage.Send",
  "Tasks.ReadWrite",
  "Notes.ReadWrite",
];

// Other people's mailboxes and calendars the user has been given access to.
// Calendars.Read.Shared is already a read scope (free/busy needs it).
const SHARED_READ_SCOPES = ["Mail.Read.Shared"];
const SHARED_WRITE_SCOPES = ["Mail.ReadWrite.Shared", "Mail.Send.Shared", "Calendars.ReadWrite.Shared"];

// Only an admin can grant these: requesting one unapproved fails the whole
// sign-in, so they are opt-in (MCP_ENABLE_ADMIN_SCOPES).
const ADMIN_SCOPES = ["Place.Read.All", "User.Read.All", "OnlineMeetings.Read", "OnlineMeetingTranscript.Read.All"];

/** Microsoft's "Microsoft Graph Command Line Tools" public client.
 *  Convenient for trying the server out; override MCP_CLIENT_ID for production
 *  so you get clean audit logs and can scope your own permissions. */
export const DEFAULT_PUBLIC_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";

/** Shortest MCP_TOKEN_CACHE_KEY accepted; scrypt stretches it, but cannot add entropy. */
export const MIN_CACHE_KEY_LENGTH = 16;

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
    .option("--client-certificate <path>", "PEM certificate with private key (client-credentials only)")
    .option("--user <upn>", "The user app-only mode acts for (client-credentials only)")
    .option("--redirect-uri <url>", "Redirect URI (interactive only)")
    .option("--scopes <csv>", "Comma-separated Graph scopes (overrides default)")
    .option("--enable-writes", "Enable all mutating tools")
    .option("--disabled-tools <csv>", "Comma-separated tool names to hide")
    .option("--token-cache <path>", "Token cache file path")
    .option("--list-tools", "Print the tool catalogue and exit (no server, no auth)")
    .option("--logout", "Remove this app's stored tokens and sign-in record, then exit (forces re-auth on next run)")
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
  const clientCertificatePath = opts.clientCertificate ?? process.env.MCP_CLIENT_CERTIFICATE_PATH;
  const user = opts.user ?? process.env.MCP_USER;
  if (authMode === "client-credentials") {
    if (!user) {
      throw new Error(
        "client-credentials mode has no signed-in user: set MCP_USER to the user (or shared mailbox) " +
          "whose mail, calendar and files the server works with.",
      );
    }
    if (["common", "organizations", "consumers"].includes(tenantId)) {
      throw new Error(`client-credentials mode needs your tenant id or domain in MCP_TENANT_ID, not "${tenantId}".`);
    }
  }
  // Shared mailboxes are reached with the .Shared scopes; opt-in, since new scopes need new consent.
  const sharedMailboxes = authMode !== "client-credentials" && envBool("MCP_ENABLE_SHARED_MAILBOXES") === true;
  // App-only permissions are whatever the admin granted; the flag is for signed-in users.
  const adminScopes = authMode !== "client-credentials" && envBool("MCP_ENABLE_ADMIN_SCOPES") === true;
  const redirectUri = opts.redirectUri ?? process.env.MCP_REDIRECT_URI;

  const enableWrites = Boolean(opts.enableWrites) || envBool("MCP_ENABLE_WRITES") === true;

  const perSurfaceWrites: Record<string, boolean> = {};
  for (const s of SURFACES) {
    const envName = `MCP_ENABLE_${s.toUpperCase()}_WRITE`;
    const v = envBool(envName);
    if (v !== undefined) perSurfaceWrites[s] = v;
  }

  const wantsWrites = enableWrites || Object.values(perSurfaceWrites).some(Boolean);
  const defaultScopes = [
    ...READ_SCOPES,
    ...(wantsWrites ? WRITE_SCOPES : []),
    ...(sharedMailboxes ? SHARED_READ_SCOPES : []),
    ...(sharedMailboxes && wantsWrites ? SHARED_WRITE_SCOPES : []),
    ...(adminScopes ? ADMIN_SCOPES : []),
  ];
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

  // The MCP SDK's stdio default is 10 MiB, which caps base64 uploads near 7.5 MB,
  // and a message over the limit closes the transport and ends the server.
  const maxMessageMb = Number(process.env.MCP_MAX_MESSAGE_MB ?? 64);
  if (!Number.isInteger(maxMessageMb) || maxMessageMb < 1) {
    throw new Error(`MCP_MAX_MESSAGE_MB must be a positive whole number of MiB, got "${process.env.MCP_MAX_MESSAGE_MB}".`);
  }

  // A flag would put the secret in the process list; the environment only.
  const tokenCacheKey = process.env.MCP_TOKEN_CACHE_KEY || undefined;
  if (tokenCacheKey !== undefined && tokenCacheKey.length < MIN_CACHE_KEY_LENGTH) {
    throw new Error(
      `MCP_TOKEN_CACHE_KEY must be at least ${MIN_CACHE_KEY_LENGTH} characters; generate one with: openssl rand -base64 32`,
    );
  }

  return {
    authMode,
    tenantId,
    clientId,
    clientSecret,
    clientCertificatePath,
    clientCertificatePassword: process.env.MCP_CLIENT_CERTIFICATE_PASSWORD || undefined,
    user,
    sharedMailboxes,
    adminScopes,
    redirectUri,
    tokenCachePath: opts.tokenCache ?? process.env.MCP_TOKEN_CACHE_PATH ?? defaultCachePath(),
    tokenCacheKey,
    scopes: effectiveScopes,
    enableWrites,
    perSurfaceWrites,
    disabledTools,
    logLevel: process.env.MCP_LOG_LEVEL ?? "info",
    listTools: Boolean(opts.listTools),
    logout: Boolean(opts.logout),
    maxMessageBytes: maxMessageMb * 1024 * 1024,
    downloadDir: process.env.MCP_DOWNLOAD_DIR ? path.resolve(process.env.MCP_DOWNLOAD_DIR) : undefined,
    uploadDir: process.env.MCP_UPLOAD_DIR ? path.resolve(process.env.MCP_UPLOAD_DIR) : undefined,
  };
}
