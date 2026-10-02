import { z } from "zod";
import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { batchGet } from "../../graph/batch.js";
import { escapeHtml, textToHtml } from "../../util/html.js";

export const MentionsInput = z
  .array(z.string().min(3).describe("Email address or directory id."))
  .max(20)
  .optional()
  .describe("People to @mention. Write @Name in the body where each should appear; otherwise they are put at the start.");

export const MessageBody = {
  body: z.string(),
  contentType: z.enum(["text", "html"]).default("text"),
  mentions: MentionsInput,
};

export interface ChatMessageBody {
  body: { contentType: "text" | "html"; content: string };
  mentions?: Array<{ id: number; mentionText: string; mentioned: { user: { id: string; displayName: string; userIdentityType: "aadUser" } } }>;
}

/**
 * The body Teams wants for a message, with @mentions resolved: each person is
 * looked up in the directory (one batch), and `@Their Name` in the text becomes
 * a real mention that notifies them. Text is escaped, since a mention makes the
 * message HTML.
 */
export async function messageBody(
  graph: GraphClient,
  input: { body: string; contentType: "text" | "html"; mentions?: string[] },
): Promise<ChatMessageBody> {
  if (!input.mentions?.length) return { body: { contentType: input.contentType, content: input.body } };

  const found = await batchGet(graph, input.mentions.map((u, i) => ({ id: String(i), url: `/users/${encodeURIComponent(u)}?$select=id,displayName` })));
  const people = found.map((r, i) => {
    const b = r.body as { id?: string; displayName?: string } | undefined;
    if (!b?.id) throw new Error(`Cannot mention ${input.mentions![i]}: ${r.error?.message ?? "not found in the directory"}`);
    return { id: b.id, name: b.displayName ?? input.mentions![i]! };
  });

  let html = input.contentType === "html" ? input.body : textToHtml(input.body);
  const leading: string[] = [];
  people.forEach((p, i) => {
    const tag = `<at id="${i}">${escapeHtml(p.name)}</at>`;
    const written = `@${escapeHtml(p.name)}`;
    if (html.includes(written)) html = html.replace(written, tag);
    else leading.push(tag);
  });
  return {
    body: { contentType: "html", content: (leading.length ? `${leading.join(" ")} ` : "") + html },
    mentions: people.map((p, i) => ({ id: i, mentionText: p.name, mentioned: { user: { id: p.id, displayName: p.name, userIdentityType: "aadUser" } } })),
  };
}
