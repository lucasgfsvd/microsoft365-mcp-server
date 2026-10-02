import { z } from "zod";
import type { ToolContext, ToolDefinition } from "../../types.js";
import { trimEmpty } from "../../util/schema.js";
import { AttachmentInput, attachFiles } from "../../graph/mailAttachments.js";
import { Recipient, toRecipient } from "./shared.js";
import { draftFrom, draftSummary } from "./replyDraft.js";

/** Writing mail: new messages, drafts, replies, forwards, and deleting. */

const Attachments = z
  .array(AttachmentInput)
  .max(20)
  .optional()
  .describe("Files to attach: from MCP_UPLOAD_DIR (localPath), OneDrive/SharePoint (driveItemId) or inline (contentBase64 + name). Up to 150 MB each.");

const Body = {
  subject: z.string(),
  body: z.string(),
  bodyType: z.enum(["Text", "HTML"]).default("Text"),
};

type Message = { subject: string; body: string; bodyType: string; to: Array<z.infer<typeof Recipient>>; cc?: Array<z.infer<typeof Recipient>>; bcc?: Array<z.infer<typeof Recipient>> };

const messageJson = (m: Message) =>
  trimEmpty({
    subject: m.subject,
    body: { contentType: m.bodyType, content: m.body },
    toRecipients: m.to.map(toRecipient),
    ccRecipients: m.cc?.map(toRecipient),
    bccRecipients: m.bcc?.map(toRecipient),
  });

/** A draft with its attachments; removed again if any attachment fails. */
async function draftWithAttachments(ctx: ToolContext, m: Message, attachments: AttachmentInput[]) {
  const draft = (await ctx.graph.api(`/me/messages`).post(messageJson(m))) as Record<string, unknown> & { id: string };
  try {
    const attached = await attachFiles(ctx.graph, `/me/messages/${draft.id}`, attachments, ctx);
    return { draft, attached };
  } catch (err) {
    await ctx.graph.api(`/me/messages/${draft.id}`).delete().catch(() => undefined);
    throw err;
  }
}

/** Send a draft; if sending fails, the draft is removed so no half-sent copy lingers. */
async function sendDraft(ctx: ToolContext, id: string, discardOnFailure: boolean): Promise<void> {
  try {
    await ctx.graph.api(`/me/messages/${id}/send`).post({});
  } catch (err) {
    if (discardOnFailure) await ctx.graph.api(`/me/messages/${id}`).delete().catch(() => undefined);
    throw err;
  }
}

export const composeTools: ToolDefinition[] = [
  {
    name: "mail_send_message",
    surface: "mail",
    description:
      "Send a new email, optionally with attachments. With attachments it is built as a draft, the files " +
      "added, and then sent; it is always saved to Sent Items then.",
    mutating: true,
    requiredScopes: ["Mail.Send"],
    inputSchema: z.object({
      to: z.array(Recipient).min(1),
      cc: z.array(Recipient).optional(),
      bcc: z.array(Recipient).optional(),
      ...Body,
      saveToSentItems: z.boolean().default(true),
      attachments: Attachments,
    }),
    handler: async (input, ctx) => {
      if (!input.attachments?.length) {
        await ctx.graph.api(`/me/sendMail`).post({ message: messageJson(input), saveToSentItems: input.saveToSentItems });
        return { ok: true };
      }
      const { draft, attached } = await draftWithAttachments(ctx, input, input.attachments);
      await sendDraft(ctx, draft.id, true);
      return { ok: true, attached };
    },
  },
  {
    name: "mail_create_draft",
    surface: "mail",
    description: "Create a draft message, optionally with attachments (does not send).",
    mutating: true,
    requiredScopes: ["Mail.ReadWrite"],
    inputSchema: z.object({
      to: z.array(Recipient).min(1),
      cc: z.array(Recipient).optional(),
      bcc: z.array(Recipient).optional(),
      ...Body,
      attachments: Attachments,
    }),
    handler: async (input, ctx) => {
      if (!input.attachments?.length) return ctx.graph.api(`/me/messages`).post(messageJson(input));
      const { draft, attached } = await draftWithAttachments(ctx, input, input.attachments);
      return { ...draftSummary(draft), attached };
    },
  },
  {
    name: "mail_send_draft",
    surface: "mail",
    description: "Send a draft that is already in Drafts, such as one made by mail_create_draft or mail_create_reply_draft.",
    mutating: true,
    requiredScopes: ["Mail.Send"],
    inputSchema: z.object({ id: z.string().describe("The draft's message id.") }),
    handler: async ({ id }, ctx) => {
      await sendDraft(ctx, id, false);
      return { ok: true };
    },
  },
  {
    name: "mail_reply_message",
    surface: "mail",
    description:
      "Reply to a message (or reply all) and send it at once. Your text goes above the quoted original, " +
      "as written. To let the user review first, use mail_create_reply_draft.",
    mutating: true,
    requiredScopes: ["Mail.Send"],
    inputSchema: z.object({
      id: z.string(),
      comment: z.string().describe("The reply, as plain text; line breaks are kept."),
      replyAll: z.boolean().default(false),
    }),
    handler: async ({ id, comment, replyAll }, ctx) => {
      const draft = await draftFrom(ctx.graph, id, replyAll ? "createReplyAll" : "createReply", comment);
      await sendDraft(ctx, draft.id, true);
      return { ok: true };
    },
  },
  {
    name: "mail_forward_message",
    surface: "mail",
    description:
      "Forward a message, with its attachments, to new recipients, adding your text above the original. " +
      "Sends at once, unless draft: true leaves it in Drafts for review.",
    mutating: true,
    requiredScopes: ["Mail.Send"],
    inputSchema: z.object({
      id: z.string(),
      to: z.array(Recipient).min(1),
      cc: z.array(Recipient).optional(),
      comment: z.string().optional().describe("Plain text above the forwarded message; line breaks are kept."),
      draft: z.boolean().default(false),
    }),
    handler: async (input, ctx) => {
      const fwd = await draftFrom(ctx.graph, input.id, "createForward", input.comment, {
        toRecipients: input.to.map(toRecipient),
        ...(input.cc ? { ccRecipients: input.cc.map(toRecipient) } : {}),
      });
      if (input.draft) return draftSummary(fwd);
      await sendDraft(ctx, fwd.id, true);
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
