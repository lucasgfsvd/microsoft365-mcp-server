import { z } from "zod";
import type { ToolDefinition } from "../../types.js";

interface ItemBody {
  contentType?: string;
  content?: string;
}

/** What a caller needs to find a new draft again; the quoted original stays out of the result. */
const DRAFT_FIELDS = ["id", "subject", "toRecipients", "ccRecipients", "conversationId", "isDraft", "webLink"] as const;

const escapeHtml = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/**
 * The reply body with the user's plain text put above the quoted original,
 * written for the body's own format. Graph's `comment` parameter would do
 * this, but it reads the comment as HTML: line breaks collapsed and markup
 * passed through (seen live).
 */
export function withReplyText(body: ItemBody, text: string): ItemBody {
  const content = body.content ?? "";
  if (body.contentType?.toLowerCase() !== "html") return { contentType: "text", content: `${text}\n\n${content}` };
  const html = `<div>${escapeHtml(text).replace(/\r?\n/g, "<br>")}</div>`;
  const open = /<body[^>]*>/i.exec(content);
  if (!open) return { contentType: "html", content: html + content };
  const at = open.index + open[0].length;
  return { contentType: "html", content: content.slice(0, at) + html + content.slice(at) };
}

export const replyDraftTools: ToolDefinition[] = [
  {
    name: "mail_create_reply_draft",
    surface: "mail",
    description:
      "Draft a reply inside the message's conversation, without sending it. Outlook fills in the " +
      'recipients, the "RE:" subject and the quoted original; your text goes above the quote. The ' +
      "draft waits in Drafts for the user to review and send.",
    mutating: true,
    requiredScopes: ["Mail.ReadWrite"],
    inputSchema: z.object({
      id: z.string().describe("The message being replied to."),
      comment: z.string().describe("The reply, as plain text; line breaks are kept. Placed above the quoted original."),
      replyAll: z.boolean().default(false),
    }),
    handler: async ({ id, comment, replyAll }, ctx) => {
      const action = replyAll ? "createReplyAll" : "createReply";
      let draft = (await ctx.graph.api(`/me/messages/${id}/${action}`).post({})) as Record<string, unknown> & {
        id: string;
        body?: ItemBody;
      };
      if (comment) {
        try {
          const updated: typeof draft | undefined = await ctx.graph
            .api(`/me/messages/${draft.id}`)
            .patch({ body: withReplyText(draft.body ?? {}, comment) });
          draft = updated ?? draft;
        } catch (err) {
          // Do not leave an empty reply behind in Drafts.
          await ctx.graph.api(`/me/messages/${draft.id}`).delete().catch(() => undefined);
          throw err;
        }
      }
      return Object.fromEntries(DRAFT_FIELDS.filter((k) => k in draft).map((k) => [k, draft[k]]));
    },
  },
];
