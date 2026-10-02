import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { DrivePath, Filename } from "../../util/schema.js";
import { downloadInline, downloadToDir } from "../../graph/download.js";
import { uploadContent } from "../../graph/upload.js";
import { drivePrefix, itemByPath, ScopeInput } from "./scope.js";
import { downloadOptions, uploadOptions } from "./transfer.js";

interface ItemRef {
  id: string;
  name: string;
  size?: number;
  parentReference?: { path?: string };
}

/** The folder path of an item's parent, as the drive-relative path tools take ("/Reports"). */
export function parentPathOf(item: ItemRef): string {
  const p = item.parentReference?.path ?? "";
  const i = p.indexOf("root:");
  return i < 0 ? "/" : decodeURIComponent(p.slice(i + 5)) || "/";
}

/** Formats OneDrive can turn into PDF (Graph's ?format=pdf). */
const PDF_SOURCES = /\.(docx?|dotx?|dotm|pptx?|ppsx?|xlsx?|xlsm|odt|odp|ods|rtf|html?|eml|msg|epub|md|markdown|tiff?)$/i;

export const fileManageTools: ToolDefinition[] = [
  {
    name: "files_move",
    surface: "files",
    description: "Move a file or folder to another folder, rename it, or both. Fails rather than replace an item of the same name.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: ScopeInput.extend({
      itemId: z.string(),
      destinationParentPath: DrivePath.optional().describe("Folder to move into, e.g. '/Archive/2025'."),
      newName: Filename.optional(),
    }).refine((m) => m.destinationParentPath !== undefined || m.newName !== undefined, { message: "Give destinationParentPath, newName, or both" }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      const patch: Record<string, unknown> = { "@microsoft.graph.conflictBehavior": "fail" };
      if (input.newName) patch.name = input.newName;
      if (input.destinationParentPath) {
        const folder = (await ctx.graph.api(itemByPath(base, input.destinationParentPath)).select("id,folder").get()) as { id: string; folder?: unknown };
        if (!folder.folder) throw new Error(`${input.destinationParentPath} is not a folder.`);
        patch.parentReference = { id: folder.id };
      }
      const moved = (await ctx.graph.api(`${base}/items/${input.itemId}`).patch(patch)) as ItemRef & { webUrl?: string };
      return { id: moved.id, name: moved.name, folder: parentPathOf(moved), webUrl: moved.webUrl };
    },
  },
  {
    name: "files_export_pdf",
    surface: "files",
    description:
      "Turn a Word, PowerPoint, Excel (or other Office, HTML, Markdown, email) file into a PDF, saved next to " +
      "the original in the drive (same name, .pdf, replacing an older export), or to MCP_DOWNLOAD_DIR with " +
      "saveTo: 'disk'. Attach the result to mail with its driveItemId.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: ScopeInput.extend({ itemId: z.string(), saveTo: z.enum(["drive", "disk"]).default("drive") }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      const item = (await ctx.graph.api(`${base}/items/${input.itemId}`).select("id,name,size,parentReference").get()) as ItemRef;
      if (!PDF_SOURCES.test(item.name)) throw new Error(`${item.name} cannot be converted to PDF by OneDrive.`);
      const pdfName = item.name.replace(/\.[^.]+$/, "") + ".pdf";
      const source = `${base}/items/${input.itemId}/content?format=pdf`;
      if (input.saveTo === "disk") {
        if (!ctx.config.downloadDir) throw new Error("Saving to disk is off: set MCP_DOWNLOAD_DIR to the folder the server may write downloads into.");
        return { name: pdfName, ...(await downloadToDir(ctx.graph, source, ctx.config.downloadDir, pdfName, downloadOptions(ctx, pdfName, item.size ?? 0))) };
      }
      const pdf = await downloadInline(ctx.graph, source);
      const saved = (await uploadContent(ctx.graph, base, { parentPath: parentPathOf(item), filename: pdfName }, pdf, uploadOptions(ctx, pdfName))) as ItemRef & { webUrl?: string };
      return { driveItemId: saved.id, name: saved.name, size: saved.size, folder: parentPathOf(item), webUrl: saved.webUrl };
    },
  },
];
