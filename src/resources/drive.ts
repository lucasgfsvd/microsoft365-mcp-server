import { downloadInline, getItemMeta, INLINE_LIMIT } from "../graph/download.js";
import { convertForConversation } from "../content/convert.js";
import { recentFiles, type RecentFile } from "../graph/recentFiles.js";
import { enc, uriFor, type ResourceContent, type ResourceKind } from "./types.js";

/**
 * File content for the conversation: Word, PowerPoint, Excel and PDF as their
 * text, other text files as text, images and anything else as a blob, and
 * nothing over the inline limit, which has to fit the client's message size.
 */
export async function fileContent(uri: string, bytes: Buffer, mimeType: string | undefined, name: string): Promise<ResourceContent> {
  const c = await convertForConversation(bytes, mimeType, name);
  switch (c.kind) {
    case "text":
      return { uri, mimeType: c.mimeType, text: c.note ? `[${c.note}]\n\n${c.text}` : c.text };
    case "image":
      return { uri, mimeType: c.mimeType, blob: c.data };
    case "binary":
      return { uri, mimeType: c.mimeType, blob: c.base64 };
  }
}

export const driveResource: ResourceKind = {
  key: "drive",
  requiresTool: "files_download",
  template: {
    uriTemplate: "m365://drive/{driveId}/{itemId}",
    name: "drive-file",
    title: "OneDrive / SharePoint file",
    description: `A file's content: Word, PowerPoint and PDF as text, Excel as CSV per sheet, other text files as text, up to ${INLINE_LIMIT / 1024 / 1024} MB.`,
  },
  // Microsoft retires /me/drive/recent after November 2026; search replaces it.
  recent: { fetch: (graph) => recentFiles(graph, 20) },
  toEntries: (body) =>
    ((body ?? []) as RecentFile[]).flatMap((f) =>
      f.driveId
        ? [{
            uri: uriFor("drive", f.driveId, f.id),
            name: `file: ${f.name ?? f.id}`,
            description: f.lastModifiedDateTime ? `Modified ${f.lastModifiedDateTime}${f.lastModifiedBy ? ` by ${f.lastModifiedBy}` : ""}` : undefined,
          }]
        : [],
    ),
  async read(graph, [driveId, itemId], uri) {
    if (!itemId) throw new Error("Drive URIs look like m365://drive/{driveId}/{itemId}.");
    const itemUrl = `/drives/${enc(driveId!)}/items/${enc(itemId)}`;
    const meta = await getItemMeta(graph, itemUrl);
    if (meta.folder) throw new Error(`${meta.name} is a folder; attach a file instead.`);
    if (meta.size > INLINE_LIMIT) {
      throw new Error(
        `${meta.name} is ${(meta.size / 1024 / 1024).toFixed(1)} MB, too large to attach. ` +
          "Ask for it with files_download and saveToDisk: true instead.",
      );
    }
    return fileContent(uri, await downloadInline(graph, `${itemUrl}/content`), meta.file?.mimeType, meta.name);
  },
};
