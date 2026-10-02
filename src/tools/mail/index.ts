import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { fetchPage } from "../../graph/pagination.js";
import { PaginationInput } from "../../util/schema.js";
import { tidyText } from "../../util/text.js";
import { replyDraftTools } from "./replyDraft.js";
import { composeTools } from "./compose.js";
import { organizeTools } from "./organize.js";
import { attachmentTools } from "./attachments.js";
import { mailSettingsTools } from "./settings.js";

/** An instant to filter on: a full ISO date-time with offset, or a date (midnight UTC). */
const Instant = z.union([z.iso.datetime({ offset: true }), z.iso.date()]);
const toUtc = (s: string) => new Date(s.length === 10 ? `${s}T00:00:00Z` : s).toISOString();

/**
 * The $filter for a folder listing. Graph wants the $orderby property
 * (receivedDateTime) to lead the filter, so the date clauses come first.
 */
export function listFilter(opts: { receivedAfter?: string; receivedBefore?: string; unreadOnly?: boolean }): string | undefined {
  const parts: string[] = [];
  if (opts.receivedAfter) parts.push(`receivedDateTime ge ${toUtc(opts.receivedAfter)}`);
  if (opts.receivedBefore) parts.push(`receivedDateTime lt ${toUtc(opts.receivedBefore)}`);
  if (opts.unreadOnly) parts.push("isRead eq false");
  return parts.length ? parts.join(" and ") : undefined;
}

/**
 * Strip invisible padding from a message's plain-text body. An HTML body is
 * left alone: collapsing its whitespace could change how it renders (<pre>).
 * `bodyPreview` is tidied for every tool by serializeResult.
 */
export function tidyMessage<M extends { body?: { content?: string } }>(msg: M, bodyFormat: "text" | "html"): M {
  const out = { ...msg };
  if (bodyFormat === "text" && typeof out.body?.content === "string") out.body = { ...out.body, content: tidyText(out.body.content) };
  return out;
}

export const mailTools: ToolDefinition[] = [
  {
    name: "mail_list_messages",
    surface: "mail",
    description:
      "List messages in a folder (default: Inbox), newest first. Narrow by when they arrived with " +
      "receivedAfter (inclusive) and receivedBefore (exclusive), e.g. this week's mail.",
    requiredScopes: ["Mail.Read"],
    inputSchema: PaginationInput.extend({
      folder: z.string().default("Inbox").describe("Well-known name (Inbox, SentItems, Drafts, DeletedItems) or folder id."),
      unreadOnly: z.boolean().optional(),
      receivedAfter: Instant.optional().describe("Received at or after this, e.g. '2026-09-28' (midnight UTC) or '2026-09-28T08:00:00+02:00'."),
      receivedBefore: Instant.optional().describe("Received before this; same forms as receivedAfter."),
    }).refine((d) => !d.receivedAfter || !d.receivedBefore || toUtc(d.receivedAfter) < toUtc(d.receivedBefore), {
      message: "receivedAfter must be earlier than receivedBefore",
    }),
    handler: async (input, ctx) => {
      return fetchPage(ctx.graph, `/me/mailFolders/${input.folder}/messages`, {
        ...input,
        filter: listFilter(input),
        orderBy: "receivedDateTime desc",
        select: ["id", "subject", "from", "toRecipients", "receivedDateTime", "isRead", "bodyPreview", "hasAttachments"],
      });
    },
  },
  {
    name: "mail_search_messages",
    surface: "mail",
    description: "Full-text search across mail using Microsoft Search.",
    requiredScopes: ["Mail.Read"],
    inputSchema: PaginationInput.extend({
      query: z.string().min(1).describe("KQL or free text, e.g. 'from:alice subject:Q3'"),
    }),
    handler: async (input, ctx) =>
      fetchPage(ctx.graph, `/me/messages`, {
        ...input,
        search: input.query,
        select: ["id", "subject", "from", "receivedDateTime", "bodyPreview"],
      }),
  },
  {
    name: "mail_get_message",
    surface: "mail",
    description:
      "Fetch a single message with full body. The text body (the default) has the invisible padding " +
      "newsletters put in their preview line removed; request html for the original markup.",
    requiredScopes: ["Mail.Read"],
    inputSchema: z.object({ id: z.string(), bodyFormat: z.enum(["text", "html"]).default("text") }),
    handler: async ({ id, bodyFormat }, ctx) => {
      const msg = (await ctx.graph
        .api(`/me/messages/${id}`)
        .header("Prefer", `outlook.body-content-type="${bodyFormat}"`)
        .get()) as { body?: { contentType?: string; content?: string }; bodyPreview?: string };
      return tidyMessage(msg, bodyFormat);
    },
  },
  {
    name: "mail_list_folders",
    surface: "mail",
    description: "List mail folders.",
    requiredScopes: ["Mail.Read"],
    inputSchema: PaginationInput,
    handler: async (input, ctx) => fetchPage(ctx.graph, `/me/mailFolders`, input),
  },
  {
    name: "mail_list_attachments",
    surface: "mail",
    description: "List attachments on a message (metadata only).",
    requiredScopes: ["Mail.Read"],
    inputSchema: z.object({ messageId: z.string() }),
    handler: async ({ messageId }, ctx) => ctx.graph.api(`/me/messages/${messageId}/attachments`).get(),
  },
  ...composeTools,
  ...replyDraftTools,
  ...organizeTools,
  ...attachmentTools,
  ...mailSettingsTools,
];
