import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { trimEmpty } from "../../util/schema.js";
import { Recipient, toRecipient } from "./shared.js";

/** Writing mail: new messages, drafts, replies, and deleting. */
export const composeTools: ToolDefinition[] = [
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
