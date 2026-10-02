import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { downloadInline, downloadToDir, INLINE_LIMIT, toInline } from "../../graph/download.js";
import { convertForConversation } from "../../content/convert.js";
import { htmlToText } from "../../resources/html.js";
import { MB } from "../../util/progress.js";
import { fileOutput } from "../../util/toolOutput.js";
import { downloadOptions } from "../files/transfer.js";

interface AttachmentMeta {
  "@odata.type"?: string;
  name: string;
  contentType?: string;
  size: number;
}

interface AttachedItem {
  "@odata.type"?: string;
  subject?: string;
  from?: unknown;
  toRecipients?: unknown;
  sentDateTime?: string;
  start?: unknown;
  end?: unknown;
  body?: { contentType?: string; content?: string };
}

export const attachmentTools: ToolDefinition[] = [
  {
    name: "mail_get_attachment",
    surface: "mail",
    description:
      "Read an attachment (ids from mail_list_attachments): Word, PowerPoint, Excel and PDF as their text, " +
      `images as images you can see, an attached email or invitation as its text; up to ${MB(INLINE_LIMIT)} ` +
      "inline. With saveToDisk: true a file is saved to MCP_DOWNLOAD_DIR instead, at any size; raw: true " +
      "returns the exact bytes as base64.",
    requiredScopes: ["Mail.Read"],
    inputSchema: z.object({
      messageId: z.string(),
      attachmentId: z.string(),
      saveToDisk: z.boolean().optional(),
      raw: z.boolean().optional(),
    }),
    handler: async (input, ctx) => {
      const base = `/me/messages/${input.messageId}/attachments/${input.attachmentId}`;
      const meta = (await ctx.graph.api(base).select("id,name,contentType,size").get()) as AttachmentMeta;
      const kind = meta["@odata.type"];

      if (kind === "#microsoft.graph.referenceAttachment") {
        return { name: meta.name, note: "This attachment is a link to a cloud file, not a copy; find it with files_search." };
      }
      if (kind === "#microsoft.graph.itemAttachment") {
        // An attached email, event or contact: its fields and body as text.
        const full = (await ctx.graph.api(base).expand("microsoft.graph.itemattachment/item").get()) as { item?: AttachedItem };
        const { body, ...item } = full.item ?? {};
        const text = body?.contentType?.toLowerCase() === "html" ? htmlToText(body.content ?? "") : (body?.content ?? "");
        return { name: meta.name, item: { ...item, body: text } };
      }

      if (input.saveToDisk) {
        if (!ctx.config.downloadDir) throw new Error("Saving to disk is off: set MCP_DOWNLOAD_DIR to the folder the server may write downloads into.");
        const saved = await downloadToDir(ctx.graph, `${base}/$value`, ctx.config.downloadDir, meta.name, downloadOptions(ctx, meta.name, meta.size));
        return { name: meta.name, mimeType: meta.contentType, ...saved };
      }
      if (meta.size > INLINE_LIMIT) {
        throw new Error(`${meta.name} is ${MB(meta.size)}, over the ${MB(INLINE_LIMIT)} that can be returned inline. Call again with saveToDisk: true.`);
      }
      const bytes = await downloadInline(ctx.graph, `${base}/$value`);
      if (input.raw) return { name: meta.name, ...toInline(bytes, meta.contentType) };
      return fileOutput({ name: meta.name, byteLength: bytes.byteLength }, await convertForConversation(bytes, meta.contentType, meta.name));
    },
  },
];
