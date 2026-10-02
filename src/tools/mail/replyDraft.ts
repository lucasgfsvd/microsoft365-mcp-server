import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { textToHtml } from "../../util/html.js";

export interface ItemBody {
  contentType?: string;
  content?: string;
}

/** What a caller needs to find a new draft again; the quoted original stays out of the result. */
const DRAFT_FIELDS = ["id", "subject", "toRecipients", "ccRecipients", "conversationId", "isDraft", "webLink"] as const;

export function draftSummary(draft: Record<string, unknown>): Record<string, unknown> {
  return Object.fromEntries(DRAFT_FIELDS.filter((k) => k in draft).map((k) => [k, draft[k]]));
}

/**
 * The reply body with the user's plain text put above the quoted original,
 * written for the body's own format. Graph's `comment` parameter would do
 * this, but it reads the comment as HTML: line breaks collapsed and markup
 * passed through (seen live).
 */
export function withReplyText(body: ItemBody, text: string): ItemBody {
  const content = body.content ?? "";
  if (body.contentType?.toLowerCase() !== "html") return { contentType: "text", content: `${text}\n\n${content}` };
  const html = `<div>${textToHtml(text)}</div>`;
  const open = /<body[^>]*>/i.exec(content);
  if (!open) return { contentType: "html", content: html + content };
  const at = open.index + open[0].length;
  return { contentType: "html", content: content.slice(0, at) + html + content.slice(at) };
}

type Draft = Record<string, unknown> & { id: string; body?: ItemBody };
type Graph = Parameters<ToolDefinition["handler"]>[1]["graph"];

/**
 * Create Outlook's own reply, reply-all or forward of a message as a draft,
 * then write the user's text (and any field changes) into it. If writing fails,
 * the draft is removed rather than left behind half-made.
 */
export async function draftFrom(
  graph: Graph,
  id: string,
  action: "createReply" | "createReplyAll" | "createForward",
  text: string | undefined,
  fields: Record<string, unknown> = {},
): Promise<Draft> {
  const draft = (await graph.api(`/me/messages/${id}/${action}`).post({})) as Draft;
  const patch = { ...fields, ...(text ? { body: withReplyText(draft.body ?? {}, text) } : {}) };
  if (!Object.keys(patch).length) return draft;
  try {
    // Merged, not replaced: the id must survive an answer that omits it.
    const updated = (await graph.api(`/me/messages/${draft.id}`).patch(patch)) as Partial<Draft> | undefined;
    return { ...draft, ...(updated ?? {}), id: draft.id };
  } catch (err) {
    await graph.api(`/me/messages/${draft.id}`).delete().catch(() => undefined);
    throw err;
  }
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
    handler: async ({ id, comment, replyAll }, ctx) =>
      draftSummary(await draftFrom(ctx.graph, id, replyAll ? "createReplyAll" : "createReply", comment)),
  },
];
