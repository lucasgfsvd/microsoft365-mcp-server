import { z } from "zod";
import type { ServerConfig, Surface, ToolDefinition } from "../types.js";

/**
 * The `mailbox` argument: run a mail or calendar tool against a mailbox the
 * user was given access to (a shared mailbox, or a manager's calendar as their
 * delegate) instead of their own. Exchange checks the access; the server only
 * points the request there (see graph/targetUser.ts).
 *
 * Offered only with MCP_ENABLE_SHARED_MAILBOXES, because it needs the .Shared
 * scopes, and new scopes mean a new consent at the next sign-in.
 */

const SURFACES = new Set<Surface>(["mail", "calendar"]);

const MAILBOX_PARAM = {
  type: "string",
  description:
    "Another mailbox or calendar you have access to (a shared mailbox, or someone who made you their " +
    "delegate), by email address. Omit for your own.",
};

const Mailbox = z
  .string()
  .trim()
  .refine((s) => /^[^@\s/?#]+@[^@\s/?#]+$/.test(s), { message: "mailbox must be an email address" });

type Config = Pick<ServerConfig, "sharedMailboxes" | "authMode">;

export function offersMailbox(tool: Pick<ToolDefinition, "surface">, config: Config): boolean {
  return config.sharedMailboxes && SURFACES.has(tool.surface);
}

/** A tool's JSON schema, with the mailbox argument added. */
export function withMailboxParam<S extends { properties?: Record<string, unknown> }>(schema: S): S {
  return { ...schema, properties: { ...(schema.properties ?? {}), mailbox: MAILBOX_PARAM } };
}

/**
 * Split the mailbox argument from the tool's own. Where it is not offered it is
 * refused, never ignored: ignoring it would quietly act on the user's own mailbox.
 */
export function takeMailbox(
  tool: Pick<ToolDefinition, "name" | "surface">,
  config: Config,
  args: Record<string, unknown>,
): { mailbox?: string; args: Record<string, unknown> } {
  if (args.mailbox === undefined) return { args };
  const { mailbox, ...rest } = args;
  if (!offersMailbox(tool, config)) {
    throw new Error(
      config.authMode === "client-credentials"
        ? "App-only mode works for MCP_USER only; another mailbox needs a server of its own."
        : config.sharedMailboxes
          ? `${tool.name} works on your own data only; it takes no mailbox.`
          : "Other mailboxes are off: set MCP_ENABLE_SHARED_MAILBOXES=true, then sign in again to consent to them.",
    );
  }
  return { mailbox: Mailbox.parse(mailbox), args: rest };
}
