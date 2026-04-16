import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { fetchPage } from "../../graph/pagination.js";
import { PaginationInput, trimEmpty } from "../../util/schema.js";

const Recipient = z.object({
  address: z.string().email(),
  name: z.string().optional(),
});

const toRecipient = (r: z.infer<typeof Recipient>) => ({
  emailAddress: { address: r.address, name: r.name },
});

export const mailTools: ToolDefinition[] = [
  {
    name: "mail_list_messages",
    surface: "mail",
    description: "List messages in a folder (default: Inbox), newest first.",
    requiredScopes: ["Mail.Read"],
    inputSchema: PaginationInput.extend({
      folder: z.string().default("Inbox").describe("Well-known name (Inbox, SentItems, Drafts, DeletedItems) or folder id."),
      unreadOnly: z.boolean().optional(),
    }),
    handler: async (input, ctx) => {
      const filter = input.unreadOnly ? "isRead eq false" : undefined;
      return fetchPage(ctx.graph, `/me/mailFolders/${input.folder}/messages`, {
        ...input,
        filter,
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
    description: "Fetch a single message with full body.",
    requiredScopes: ["Mail.Read"],
    inputSchema: z.object({ id: z.string(), bodyFormat: z.enum(["text", "html"]).default("text") }),
    handler: async ({ id, bodyFormat }, ctx) =>
      ctx.graph.api(`/me/messages/${id}`).header("Prefer", `outlook.body-content-type="${bodyFormat}"`).get(),
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
];
