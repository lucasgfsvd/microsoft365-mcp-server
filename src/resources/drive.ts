import { downloadInline, getItemMeta, INLINE_LIMIT } from "../graph/download.js";
import { convertForConversation } from "../content/convert.js";
import { enc, uriFor, values, type ResourceContent, type ResourceKind } from "./types.js";

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

interface RecentItem {
  id?: string;
  name?: string;
  file?: unknown;
  lastModifiedDateTime?: string;
  parentReference?: { driveId?: string };
  // Items shared from someone else's drive are described here instead.
  remoteItem?: { id?: string; name?: string; file?: unknown; parentReference?: { driveId?: string } };
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
  recent: { id: "drive", url: "/me/drive/recent?$top=20" },
  toEntries: (body) =>
    (values(body) as RecentItem[]).flatMap((it) => {
      const item = it.remoteItem ?? it;
      const driveId = item.parentReference?.driveId ?? it.parentReference?.driveId;
      if (!item.file || !driveId || !item.id) return [];
      return [{
        uri: uriFor("drive", driveId, item.id),
        name: `file: ${item.name ?? it.name ?? item.id}`,
        description: it.lastModifiedDateTime ? `Modified ${it.lastModifiedDateTime}` : undefined,
      }];
    }),
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
