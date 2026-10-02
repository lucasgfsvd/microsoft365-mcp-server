import { downloadInline, getItemMeta, INLINE_LIMIT, toInline } from "../graph/download.js";
import { extractText as docxText } from "../ooxml/docx.js";
import { extractAllSlides } from "../ooxml/pptx.js";
import { workbookText } from "../ooxml/xlsxText.js";
import { enc, uriFor, values, type ResourceContent, type ResourceKind } from "./types.js";

const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const XLSX = /^application\/(vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet|vnd\.ms-excel\.sheet\.macroEnabled\.12)$/;

/**
 * File content for the conversation. Word, PowerPoint and Excel become their
 * text (a workbook as CSV per sheet): base64 of a zip is useless as context.
 * Other text files come back as text, anything else as a blob, and nothing over
 * the inline limit, which has to fit the client's message size.
 */
export async function fileContent(uri: string, bytes: Buffer, mimeType: string | undefined, name: string): Promise<ResourceContent> {
  if (mimeType === DOCX || /\.docx$/i.test(name)) return { uri, mimeType: "text/plain", text: docxText(bytes) };
  if ((mimeType && XLSX.test(mimeType)) || /\.xls[xm]$/i.test(name)) return { uri, mimeType: "text/plain", text: await workbookText(bytes) };
  if (mimeType === PPTX || /\.pptx$/i.test(name)) {
    const text = extractAllSlides(bytes).map((s) => `--- Slide ${s.index} ---\n${s.text}`).join("\n\n");
    return { uri, mimeType: "text/plain", text };
  }
  const inline = toInline(bytes, mimeType);
  return inline.encoding === "utf8"
    ? { uri, mimeType: mimeType ?? "text/plain", text: inline.text }
    : { uri, mimeType: mimeType ?? "application/octet-stream", blob: inline.base64 };
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
    description: `A file's content: Word and PowerPoint as text, Excel as CSV per sheet, other text files as text, up to ${INLINE_LIMIT / 1024 / 1024} MB.`,
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
