import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { downloadInline, downloadToDir, getItemMeta, INLINE_LIMIT, toInline } from "../../graph/download.js";
import { drivePrefix, ScopeInput } from "./scope.js";
import { downloadOptions } from "./transfer.js";
import { MB } from "../../util/progress.js";
import { convertForConversation } from "../../content/convert.js";
import { fileOutput } from "../../util/toolOutput.js";

export const filesDownloadTools: ToolDefinition[] = [
  {
    name: "files_download",
    surface: "files",
    description:
      "Download a file. By default the content comes back in the result, made readable: Word, " +
      "PowerPoint, Excel and PDF as their text, images as images you can see, other text files as " +
      `text, anything else as base64; only up to ${MB(INLINE_LIMIT)}, because everything returned ` +
      "lands in the conversation. `raw: true` returns the exact bytes as base64 instead.\n\n" +
      "With `saveToDisk: true` the file is streamed into the server's download folder " +
      "(MCP_DOWNLOAD_DIR) instead, at any size, and the result is just its local path, size and " +
      "SHA-256. Existing files there are never overwritten; a taken name gets a numbered suffix. " +
      "Use that for anything large or binary you do not need to read.",
    requiredScopes: ["Files.Read.All", "Sites.Read.All"],
    inputSchema: ScopeInput.extend({
      itemId: z.string(),
      saveToDisk: z
        .boolean()
        .optional()
        .describe("Stream to MCP_DOWNLOAD_DIR instead of returning the bytes. Required above the inline limit."),
      raw: z.boolean().optional().describe("Return the exact bytes as base64, without converting documents to text."),
    }),
    handler: async (input, ctx) => {
      const itemUrl = `${drivePrefix(input)}/items/${input.itemId}`;
      const meta = await getItemMeta(ctx.graph, itemUrl);
      if (meta.folder) throw new Error(`${meta.name} is a folder; use files_list_children.`);

      if (input.saveToDisk) {
        const dir = ctx.config.downloadDir;
        if (!dir) {
          throw new Error(
            "Saving to disk is off: set MCP_DOWNLOAD_DIR to the folder the server may write downloads into.",
          );
        }
        const saved = await downloadToDir(ctx.graph, `${itemUrl}/content`, dir, meta.name, downloadOptions(ctx, meta.name, meta.size));
        return { name: meta.name, mimeType: meta.file?.mimeType, ...saved };
      }

      if (meta.size > INLINE_LIMIT) {
        throw new Error(
          `${meta.name} is ${MB(meta.size)}, over the ${MB(INLINE_LIMIT)} that can be returned inline. ` +
            (ctx.config.downloadDir
              ? "Call again with saveToDisk: true to stream it to the download folder."
              : "Set MCP_DOWNLOAD_DIR and call again with saveToDisk: true, or open it in OneDrive."),
        );
      }
      const bytes = await downloadInline(ctx.graph, `${itemUrl}/content`);
      if (input.raw) return { name: meta.name, ...toInline(bytes, meta.file?.mimeType) };
      return fileOutput({ name: meta.name, byteLength: bytes.byteLength }, await convertForConversation(bytes, meta.file?.mimeType, meta.name));
    },
  },
];
