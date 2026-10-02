import { extractText, getDocumentProxy } from "unpdf";

/** Most text a PDF may turn into; a short file can hold a very long text. */
export const PDF_TEXT_LIMIT = 1_000_000;

/**
 * A PDF's text, page by page.
 *
 * The file may come from anyone (an attachment, a shared drive). pdf.js 6 no
 * longer evaluates code from fonts (a crafted font once ran code through
 * older versions); embedded font loading is off as well, since only text is wanted.
 */
export async function pdfText(bytes: Buffer, limit = PDF_TEXT_LIMIT): Promise<{ text: string; pages: number; truncated: boolean }> {
  const doc = await getDocumentProxy(new Uint8Array(bytes), { disableFontFace: true, useSystemFonts: false });
  try {
    const { totalPages, text } = await extractText(doc, { mergePages: false });
    const parts: string[] = [];
    let size = 0;
    let truncated = false;
    for (let i = 0; i < text.length; i++) {
      const part = `--- Page ${i + 1} ---\n${(text[i] ?? "").trim()}`;
      if (size + part.length > limit) {
        truncated = true;
        break;
      }
      parts.push(part);
      size += part.length + 2;
    }
    return { text: parts.join("\n\n"), pages: totalPages, truncated };
  } finally {
    await doc.loadingTask.destroy();
  }
}
