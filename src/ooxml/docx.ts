import PizZip from "pizzip";
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from "docx";
import { decodeXmlEntities, escapeRegex, escapeXml } from "./escape.js";

/** A high-level block in a generated document. The same type drives both the
 *  `docx`-library path (used for brand-new documents) and the OOXML splice path
 *  (used for appending to existing documents). */
export type DocxBlock = {
  kind: "title" | "heading1" | "heading2" | "heading3" | "paragraph" | "bullet";
  text: string;
};

/** Open the docx zip, mutate `word/document.xml`, and return the new bytes. */
function mutateDocumentXml(buf: Buffer, mutate: (xml: string) => string): Buffer {
  const zip = new PizZip(buf);
  const file = zip.file("word/document.xml");
  if (!file) throw new Error("word/document.xml not found in .docx");
  zip.file("word/document.xml", mutate(file.asText()));
  return zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
}

/** Extract plain text from a .docx, paragraph-per-line. */
export function extractText(buf: Buffer): string {
  const zip = new PizZip(buf);
  const xml = zip.file("word/document.xml")?.asText() ?? "";
  const paragraphs = xml.split(/<\/w:p>/g).map((p) => {
    const texts = [...p.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1] ?? "");
    return decodeXmlEntities(texts.join(""));
  });
  return paragraphs.filter(Boolean).join("\n");
}

/** Return one string per paragraph (preserving order, including empty paragraphs). */
export function listParagraphs(buf: Buffer): string[] {
  return extractText(buf).split("\n");
}

/** Find-and-replace inside each `<w:t>` run. Returns the new buffer and the
 *  number of literal-string occurrences replaced. Note that text broken across
 *  runs (e.g., spell-check splits) won't match. */
export function replaceText(
  buf: Buffer,
  replacements: ReadonlyArray<{ find: string; replace: string }>,
): { buf: Buffer; count: number } {
  let count = 0;
  const out = mutateDocumentXml(buf, (xml) => {
    let next = xml;
    for (const { find, replace } of replacements) {
      next = next.replace(/<w:t[^>]*>([^<]*)<\/w:t>/g, (full, inner: string) => {
        if (!inner.includes(find)) return full;
        count += (inner.match(new RegExp(escapeRegex(find), "g")) ?? []).length;
        return full.replace(inner, inner.split(find).join(escapeXml(replace)));
      });
    }
    return next;
  });
  return { buf: out, count };
}

/** Append a plain-text paragraph to the end of `<w:body>`. */
export function appendParagraph(buf: Buffer, text: string): Buffer {
  return mutateDocumentXml(buf, (xml) => {
    const para = `<w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
    return xml.replace("</w:body>", `${para}</w:body>`);
  });
}

/** Append a heading (level 1-3) referencing the document's built-in Heading style. */
export function appendHeading(buf: Buffer, text: string, level: 1 | 2 | 3): Buffer {
  const style = `Heading${level}`;
  return mutateDocumentXml(buf, (xml) => {
    const para =
      `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>` +
      `<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
    return xml.replace("</w:body>", `${para}</w:body>`);
  });
}

/** Append a bulleted list. Uses `pStyle=ListParagraph` plus `numId=1`, which
 *  exists in any document that has ever contained a list. */
export function appendBullets(buf: Buffer, items: ReadonlyArray<string>): Buffer {
  return mutateDocumentXml(buf, (xml) => {
    const paras = items
      .map(
        (item) =>
          `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>` +
          `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
          `<w:r><w:t xml:space="preserve">${escapeXml(item)}</w:t></w:r></w:p>`,
      )
      .join("");
    return xml.replace("</w:body>", `${paras}</w:body>`);
  });
}

/** Insert a plain-text paragraph after the Nth existing paragraph (0 = top). */
export function insertParagraphAt(buf: Buffer, text: string, after: number): Buffer {
  return mutateDocumentXml(buf, (xml) => {
    const para = `<w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
    if (after === 0) {
      // Splice in just after the <w:body> opening tag.
      return xml.replace(/(<w:body[^>]*>)/, `$1${para}`);
    }
    // Splitting on </w:p> gives [pre-para1, "</w:p>", pre-para2, "</w:p>", ...].
    // The slot just after the Nth </w:p> is index 2N.
    const parts = xml.split(/(<\/w:p>)/g);
    const insertAt = after * 2;
    if (insertAt > parts.length) {
      return xml.replace("</w:body>", `${para}</w:body>`);
    }
    parts.splice(insertAt, 0, para);
    return parts.join("");
  });
}

/** Delete the Nth paragraph (1-based). */
export function deleteParagraph(buf: Buffer, paragraphIndex: number): Buffer {
  return mutateDocumentXml(buf, (xml) => {
    const paraRe = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g;
    let i = 0;
    return xml.replace(paraRe, (match) => {
      i += 1;
      return i === paragraphIndex ? "" : match;
    });
  });
}

/** Render a block as a raw OOXML paragraph string (used when splicing into
 *  existing documents). The referenced styles must exist in the document's
 *  styles.xml — this is true for any .docx Word has ever created and for
 *  documents emitted by `createFromBlocks`. */
export function renderBlockXml(b: DocxBlock): string {
  const text = escapeXml(b.text);
  const run = `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  switch (b.kind) {
    case "title":
      return `<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr>${run}</w:p>`;
    case "heading1":
      return `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${run}</w:p>`;
    case "heading2":
      return `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>${run}</w:p>`;
    case "heading3":
      return `<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr>${run}</w:p>`;
    case "bullet":
      return (
        `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>` +
        `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>${run}</w:p>`
      );
    case "paragraph":
    default:
      return `<w:p>${run}</w:p>`;
  }
}

/** Build a brand-new `.docx` from a list of blocks. Uses the `docx` library
 *  (so the result has a valid styles.xml + numbering.xml — unlike raw splices). */
export async function createFromBlocks(blocks: ReadonlyArray<DocxBlock>): Promise<Buffer> {
  const children = blocks.map((b) => blockToParagraph(b));
  const doc = new Document({ sections: [{ properties: {}, children }] });
  const out = await Packer.toBuffer(doc);
  return Buffer.isBuffer(out) ? out : Buffer.from(out);
}

function blockToParagraph(b: DocxBlock): Paragraph {
  switch (b.kind) {
    case "title":
      return new Paragraph({ text: b.text, heading: HeadingLevel.TITLE });
    case "heading1":
      return new Paragraph({ text: b.text, heading: HeadingLevel.HEADING_1 });
    case "heading2":
      return new Paragraph({ text: b.text, heading: HeadingLevel.HEADING_2 });
    case "heading3":
      return new Paragraph({ text: b.text, heading: HeadingLevel.HEADING_3 });
    case "bullet":
      return new Paragraph({ children: [new TextRun(b.text)], bullet: { level: 0 } });
    case "paragraph":
    default:
      return new Paragraph({ children: [new TextRun(b.text)] });
  }
}

/** Apply template-style changes (replacements + appended blocks) to an
 *  existing template's `word/document.xml`. Used by `*_create_from_template`. */
export function applyTemplateChanges(
  buf: Buffer,
  opts: {
    replacements?: ReadonlyArray<{ find: string; replace: string }>;
    appendBlocks?: ReadonlyArray<DocxBlock>;
  },
): Buffer {
  return mutateDocumentXml(buf, (xml) => {
    let next = xml;
    if (opts.replacements?.length) {
      for (const { find, replace } of opts.replacements) {
        next = next.replace(/<w:t[^>]*>([^<]*)<\/w:t>/g, (full, inner: string) => {
          if (!inner.includes(find)) return full;
          return full.replace(inner, inner.split(find).join(escapeXml(replace)));
        });
      }
    }
    if (opts.appendBlocks?.length) {
      const extra = opts.appendBlocks.map((b) => renderBlockXml(b)).join("");
      next = next.replace("</w:body>", `${extra}</w:body>`);
    }
    return next;
  });
}
