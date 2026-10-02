import { toInline } from "../graph/download.js";
import { extractText as docxText } from "../ooxml/docx.js";
import { extractAllSlides } from "../ooxml/pptx.js";
import { workbookText } from "../ooxml/xlsxText.js";
import { pdfText } from "./pdf.js";

/**
 * A file, made into something a model can use: text it can read, an image it
 * can see, or, failing both, the bytes. Shared by downloads, mail attachments
 * and MCP resources, so a file reads the same however it arrives.
 */
export type Converted =
  | { kind: "text"; text: string; mimeType: string; note?: string }
  | { kind: "image"; data: string; mimeType: string }
  | { kind: "binary"; base64: string; mimeType: string };

const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const XLSX = /^application\/(vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet|vnd\.ms-excel\.sheet\.macroEnabled\.12)$/;
/** The image types multimodal models take. */
const IMAGES: Record<string, string> = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", webp: "image/webp" };

const ext = (name: string) => /\.([a-z0-9]+)$/i.exec(name)?.[1]?.toLowerCase() ?? "";
const asText = (text: string, note?: string): Converted => ({ kind: "text", text, mimeType: "text/plain", ...(note ? { note } : {}) });

/** Word, PowerPoint, Excel or PDF as text; undefined for any other kind of file. */
async function documentText(bytes: Buffer, type: string, e: string): Promise<Converted | undefined> {
  if (type === DOCX || e === "docx") return asText(docxText(bytes));
  if (type === PPTX || e === "pptx") return asText(extractAllSlides(bytes).map((s) => `--- Slide ${s.index} ---\n${s.text}`).join("\n\n"));
  if (XLSX.test(type) || e === "xlsx" || e === "xlsm") return asText(await workbookText(bytes));
  if (type === "application/pdf" || e === "pdf") {
    const pdf = await pdfText(bytes);
    const empty = !pdf.text.replace(/--- Page \d+ ---/g, "").trim();
    return asText(
      pdf.text,
      empty
        ? "No text layer: probably a scanned image. Fetch it with raw: true, or open it in a viewer."
        : pdf.truncated
          ? `Text cut short; the PDF has ${pdf.pages} pages.`
          : undefined,
    );
  }
  return undefined;
}

export async function convertForConversation(bytes: Buffer, mimeType: string | undefined, name: string): Promise<Converted> {
  const type = mimeType?.toLowerCase().split(";")[0]?.trim() ?? "";
  const e = ext(name);

  // A damaged or encrypted document is still a file: hand back its bytes rather than fail.
  const doc = await documentText(bytes, type, e).catch(() => null);
  if (doc) return doc;
  if (doc === null) return { kind: "binary", base64: bytes.toString("base64"), mimeType: mimeType ?? "application/octet-stream" };

  const image = Object.values(IMAGES).includes(type) ? type : IMAGES[e];
  if (image) return { kind: "image", data: bytes.toString("base64"), mimeType: image };

  const inline = toInline(bytes, mimeType);
  return inline.encoding === "utf8"
    ? { kind: "text", text: inline.text, mimeType: mimeType ?? "text/plain" }
    : { kind: "binary", base64: inline.base64, mimeType: mimeType ?? "application/octet-stream" };
}
