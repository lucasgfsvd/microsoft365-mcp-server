import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { fetchPage } from "../../graph/pagination.js";
import { PaginationInput, trimEmpty } from "../../util/schema.js";
import { tidyText } from "../../util/text.js";
import { replyDraftTools } from "./replyDraft.js";

const Recipient = z.object({
  address: z.email(),
  name: z.string().optional(),
});

const toRecipient = (r: z.infer<typeof Recipient>) => ({
  emailAddress: { address: r.address, name: r.name },
});

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
  {
    name: "mail_send_message",
    surface: "mail",
    description: "Send a new email.",
    mutating: true,
    requiredScopes: ["Mail.Send"],
    inputSchema: z.object({
      to: z.array(Recipient).min(1),
      cc: z.array(Recipient).optional(),
      bcc: z.array(Recipient).optional(),
      subject: z.string(),
      body: z.string(),
      bodyType: z.enum(["Text", "HTML"]).default("Text"),
      saveToSentItems: z.boolean().default(true),
    }),
    handler: async (input, ctx) => {
      await ctx.graph.api(`/me/sendMail`).post({
        message: trimEmpty({
          subject: input.subject,
          body: { contentType: input.bodyType, content: input.body },
          toRecipients: input.to.map(toRecipient),
          ccRecipients: input.cc?.map(toRecipient),
          bccRecipients: input.bcc?.map(toRecipient),
        }),
        saveToSentItems: input.saveToSentItems,
      });
      return { ok: true };
    },
  },
  {
    name: "mail_create_draft",
    surface: "mail",
    description: "Create a draft message (does not send).",
    mutating: true,
    requiredScopes: ["Mail.ReadWrite"],
    inputSchema: z.object({
      to: z.array(Recipient).min(1),
      cc: z.array(Recipient).optional(),
      subject: z.string(),
      body: z.string(),
      bodyType: z.enum(["Text", "HTML"]).default("Text"),
    }),
    handler: async (input, ctx) =>
      ctx.graph.api(`/me/messages`).post({
        subject: input.subject,
        body: { contentType: input.bodyType, content: input.body },
        toRecipients: input.to.map(toRecipient),
        ccRecipients: input.cc?.map(toRecipient),
      }),
  },
  {
    name: "mail_reply_message",
    surface: "mail",
    description: "Reply to a message (reply or replyAll).",
    mutating: true,
    requiredScopes: ["Mail.Send"],
    inputSchema: z.object({
      id: z.string(),
      comment: z.string(),
      replyAll: z.boolean().default(false),
    }),
    handler: async ({ id, comment, replyAll }, ctx) => {
      const action = replyAll ? "replyAll" : "reply";
      await ctx.graph.api(`/me/messages/${id}/${action}`).post({ comment });
      return { ok: true };
    },
  },
  {
    name: "mail_delete_message",
    surface: "mail",
    description: "Move a message to Deleted Items.",
    mutating: true,
    requiredScopes: ["Mail.ReadWrite"],
    inputSchema: z.object({ id: z.string() }),
    handler: async ({ id }, ctx) => {
      await ctx.graph.api(`/me/messages/${id}`).delete();
      return { ok: true };
    },
  },
  ...replyDraftTools,
];
