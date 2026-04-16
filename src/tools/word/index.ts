import { z } from "zod";
import PizZip from "pizzip";
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from "docx";
import type { ToolDefinition, ToolContext } from "../../types.js";

/**
 * Word support: download the .docx, manipulate document.xml (OOXML) with PizZip,
 * then upload back. For creating brand-new documents we use the `docx` library.
 */
const DocRef = z.object({
  driveId: z.string().optional(),
  siteId: z.string().optional(),
  itemId: z.string(),
});

function drivePrefix(r: { driveId?: string; siteId?: string }): string {
  if (r.driveId) return `/drives/${r.driveId}`;
  if (r.siteId) return `/sites/${r.siteId}/drive`;
  return `/me/drive`;
}

async function downloadDocx(ctx: ToolContext, ref: z.infer<typeof DocRef>): Promise<Buffer> {
  const base = drivePrefix(ref);
  const stream: NodeJS.ReadableStream = await ctx.graph.api(`${base}/items/${ref.itemId}/content`).getStream();
  const chunks: Buffer[] = [];
  for await (const c of stream) chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c));
  return Buffer.concat(chunks);
}

async function uploadDocx(ctx: ToolContext, ref: z.infer<typeof DocRef>, buf: Buffer): Promise<unknown> {
  const base = drivePrefix(ref);
  return ctx.graph.api(`${base}/items/${ref.itemId}/content`).put(buf);
}

async function uploadNewDocxToPath(
  ctx: ToolContext,
  scope: { driveId?: string; siteId?: string; parentPath: string; filename: string },
  buf: Buffer,
): Promise<unknown> {
  const base = drivePrefix(scope);
  const parent = scope.parentPath.replace(/\/$/, "");
  const path = `${base}/root:${parent}/${scope.filename}:/content`;
  return ctx.graph.api(path).put(buf);
}

async function downloadTemplateBytes(
  ctx: ToolContext,
  t: {
    templateDriveId?: string;
    templateSiteId?: string;
    templateItemId?: string;
    templatePath?: string;
  },
): Promise<Buffer> {
  const base = t.templateDriveId
    ? `/drives/${t.templateDriveId}`
    : t.templateSiteId
      ? `/sites/${t.templateSiteId}/drive`
      : `/me/drive`;
  const url = t.templateItemId
    ? `${base}/items/${t.templateItemId}/content`
    : `${base}/root:${t.templatePath}:/content`;
  const stream: NodeJS.ReadableStream = await ctx.graph.api(url).getStream();
  const chunks: Buffer[] = [];
  for await (const c of stream) chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c));
  return Buffer.concat(chunks);
}

function extractTextFromDocx(buf: Buffer): string {
  const zip = new PizZip(buf);
  const xml = zip.file("word/document.xml")?.asText() ?? "";
  const paragraphs = xml.split(/<\/w:p>/g).map((p) => {
    const texts = [...p.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]);
    return texts.join("");
  });
  return paragraphs.filter(Boolean).join("\n");
}

async function mutateDocumentXml(
  ctx: ToolContext,
  ref: z.infer<typeof DocRef>,
  mutate: (doc: string) => string,
): Promise<void> {
  const buf = await downloadDocx(ctx, ref);
  const zip = new PizZip(buf);
  const file = zip.file("word/document.xml");
  if (!file) throw new Error("word/document.xml not found in .docx");
  const updated = mutate(file.asText());
  zip.file("word/document.xml", updated);
  const out = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
  await uploadDocx(ctx, ref, out);
}

const BlockSpec = z.object({
  kind: z.enum(["title", "heading1", "heading2", "heading3", "paragraph", "bullet"]),
  text: z.string(),
});

export const wordTools: ToolDefinition[] = [
  {
    name: "word_read_text",
    surface: "word",
    description: "Download a .docx and extract its plain text (OOXML parsing).",
    requiredScopes: ["Files.Read.All", "Sites.Read.All"],
    inputSchema: DocRef,
    handler: async (input, ctx) => {
      const buf = await downloadDocx(ctx, input);
      return { text: extractTextFromDocx(buf) };
    },
  },
  {
    name: "word_list_paragraphs",
    surface: "word",
    description: "Return each paragraph as a separate string (useful for targeted replacements).",
    requiredScopes: ["Files.Read.All", "Sites.Read.All"],
    inputSchema: DocRef,
    handler: async (input, ctx) => {
      const buf = await downloadDocx(ctx, input);
      return { paragraphs: extractTextFromDocx(buf).split("\n") };
    },
  },
  {
    name: "word_replace_text",
    surface: "word",
    description:
      "Find-and-replace across document.xml. Text is matched against each <w:t> run's content (so matches broken across runs may not be found).",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: DocRef.extend({
      replacements: z.array(z.object({ find: z.string().min(1), replace: z.string() })).min(1),
    }),
    handler: async (input, ctx) => {
      const buf = await downloadDocx(ctx, input);
      const zip = new PizZip(buf);
      const file = zip.file("word/document.xml");
      if (!file) throw new Error("word/document.xml not found in .docx");
      let xml = file.asText();
      let total = 0;
      for (const { find, replace } of input.replacements) {
        xml = xml.replace(/<w:t[^>]*>([^<]*)<\/w:t>/g, (full, inner: string) => {
          if (!inner.includes(find)) return full;
          total += (inner.match(new RegExp(escapeRegex(find), "g")) ?? []).length;
          return full.replace(inner, inner.split(find).join(escapeXml(replace)));
        });
      }
      zip.file("word/document.xml", xml);
      const out = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
      await uploadDocx(ctx, input, out);
      return { replacementsApplied: total };
    },
  },
  {
    name: "word_append_paragraph",
    surface: "word",
    description: "Append a plain-text paragraph to the end of a .docx.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: DocRef.extend({ text: z.string() }),
    handler: async (input, ctx) => {
      await mutateDocumentXml(ctx, input, (xml) => {
        const newPara = `<w:p><w:r><w:t xml:space="preserve">${escapeXml(input.text)}</w:t></w:r></w:p>`;
        return xml.replace("</w:body>", `${newPara}</w:body>`);
      });
      return { ok: true };
    },
  },
  {
    name: "word_create_document",
    surface: "word",
    description:
      "Create a brand new .docx from a list of block specs and upload it to OneDrive/SharePoint. " +
      "Supports title, heading1/2/3, paragraph, and bullet blocks. Uses the 'docx' library so the " +
      "output has a valid styles.xml (unlike raw OOXML appends).",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: z.object({
      driveId: z.string().optional(),
      siteId: z.string().optional(),
      parentPath: z.string().describe("Parent folder path, e.g. '/Reports'"),
      filename: z.string().describe("Filename including .docx extension"),
      blocks: z.array(BlockSpec).min(1),
    }),
    handler: async (input, ctx) => {
      const blocks = input.blocks as z.infer<typeof BlockSpec>[];
      const children = blocks.map((b) => blockToParagraph(b));
      const doc = new Document({ sections: [{ properties: {}, children }] });
      const out = await Packer.toBuffer(doc);
      const filename = input.filename.endsWith(".docx") ? input.filename : `${input.filename}.docx`;
      const result = await uploadNewDocxToPath(
        ctx,
        { driveId: input.driveId, siteId: input.siteId, parentPath: input.parentPath, filename },
        Buffer.isBuffer(out) ? out : Buffer.from(out),
      );
      return { ok: true, blocks: input.blocks.length, driveItem: result };
    },
  },
  {
    name: "word_create_from_template",
    surface: "word",
    description:
      "Create a new .docx by copying a template, applying placeholder replacements (e.g. {{CLIENT}} → 'Acme'), " +
      "and optionally appending structured blocks (headings, bullets, paragraphs). Document styles, themes, " +
      "and existing formatting from the template are preserved. Identify the template by templateItemId OR templatePath.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: z
      .object({
        templateDriveId: z.string().optional(),
        templateSiteId: z.string().optional(),
        templateItemId: z.string().optional(),
        templatePath: z.string().optional().describe("Template path, e.g. '/Templates/proposal.docx'"),
        driveId: z.string().optional(),
        siteId: z.string().optional(),
        parentPath: z.string().describe("Destination folder path"),
        filename: z.string().describe("New filename, including .docx"),
        replacements: z
          .array(z.object({ find: z.string().min(1), replace: z.string() }))
          .optional(),
        appendBlocks: z.array(BlockSpec).optional(),
      })
      .refine((d) => Boolean(d.templateItemId) || Boolean(d.templatePath), {
        message: "Provide templateItemId or templatePath",
      }),
    handler: async (input, ctx) => {
      const buf = await downloadTemplateBytes(ctx, input);
      const zip = new PizZip(buf);
      const docFile = zip.file("word/document.xml");
      if (!docFile) throw new Error("word/document.xml not found in template");
      let xml = docFile.asText();

      if (input.replacements && input.replacements.length) {
        for (const { find, replace } of input.replacements) {
          xml = xml.replace(/<w:t[^>]*>([^<]*)<\/w:t>/g, (full, inner: string) => {
            if (!inner.includes(find)) return full;
            return full.replace(inner, inner.split(find).join(escapeXml(replace)));
          });
        }
      }

      if (input.appendBlocks && input.appendBlocks.length) {
        const extra = (input.appendBlocks as z.infer<typeof BlockSpec>[])
          .map((b) => renderBlockXml(b))
          .join("");
        xml = xml.replace("</w:body>", `${extra}</w:body>`);
      }

      zip.file("word/document.xml", xml);
      const out = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
      const filename = input.filename.endsWith(".docx") ? input.filename : `${input.filename}.docx`;
      const result = await uploadNewDocxToPath(
        ctx,
        { driveId: input.driveId, siteId: input.siteId, parentPath: input.parentPath, filename },
        out,
      );
      return { ok: true, driveItem: result };
    },
  },
  {
    name: "word_append_heading",
    surface: "word",
    description:
      "Append a heading (level 1-3) to the end of a .docx. Uses the document's built-in Heading style " +
      "(references pStyle — the style itself must already exist in styles.xml, which is true for any " +
      ".docx created by Word or word_create_document).",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: DocRef.extend({
      text: z.string(),
      level: z.union([z.literal(1), z.literal(2), z.literal(3)]).default(1),
    }),
    handler: async (input, ctx) => {
      const style = `Heading${input.level}`;
      await mutateDocumentXml(ctx, input, (xml) => {
        const newPara =
          `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>` +
          `<w:r><w:t xml:space="preserve">${escapeXml(input.text)}</w:t></w:r></w:p>`;
        return xml.replace("</w:body>", `${newPara}</w:body>`);
      });
      return { ok: true, level: input.level };
    },
  },
  {
    name: "word_append_bullets",
    surface: "word",
    description:
      "Append a bulleted list to the end of a .docx. Each item becomes one paragraph styled as a list " +
      "item. Uses pStyle='ListParagraph' plus numPr numbering; the referenced numbering must exist in " +
      "numbering.xml (present in any .docx that has ever contained a list, and in documents created by " +
      "word_create_document with a bullet block).",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: DocRef.extend({ items: z.array(z.string()).min(1) }),
    handler: async (input, ctx) => {
      await mutateDocumentXml(ctx, input, (xml) => {
        const items = input.items as string[];
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
      return { ok: true, items: input.items.length };
    },
  },
  {
    name: "word_insert_paragraph_at",
    surface: "word",
    description:
      "Insert a plain-text paragraph at 1-based position `after` (0 = at the very top). " +
      "Counts <w:p> elements in document.xml.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: DocRef.extend({
      text: z.string(),
      after: z.number().int().min(0).describe("Insert after this paragraph index. 0 inserts at top."),
    }),
    handler: async (input, ctx) => {
      await mutateDocumentXml(ctx, input, (xml) => {
        const newPara = `<w:p><w:r><w:t xml:space="preserve">${escapeXml(input.text)}</w:t></w:r></w:p>`;
        // Split document.xml on </w:p> so we can inject at a specific paragraph boundary.
        const parts = xml.split(/(<\/w:p>)/g);
        // Paragraph N ends with the 2N-th element (closing </w:p>). After paragraph N is between 2N and 2N+1.
        const insertAt = input.after * 2 + 2; // slot index after Nth </w:p>
        if (insertAt > parts.length) {
          // Fall through to append at body end.
          return xml.replace("</w:body>", `${newPara}</w:body>`);
        }
        parts.splice(insertAt, 0, newPara);
        return parts.join("");
      });
      return { ok: true };
    },
  },
  {
    name: "word_delete_paragraph",
    surface: "word",
    description: "Delete a paragraph by 1-based index from an existing .docx.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: DocRef.extend({ paragraphIndex: z.number().int().min(1) }),
    handler: async (input, ctx) => {
      await mutateDocumentXml(ctx, input, (xml) => {
        // Match <w:p ...>...</w:p> greedily across the doc, replace the Nth with empty.
        const paraRe = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g;
        let i = 0;
        return xml.replace(paraRe, (match) => {
          i += 1;
          return i === input.paragraphIndex ? "" : match;
        });
      });
      return { ok: true, deletedIndex: input.paragraphIndex };
    },
  },
];

function renderBlockXml(b: z.infer<typeof BlockSpec>): string {
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

function blockToParagraph(b: z.infer<typeof BlockSpec>): Paragraph {
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
      return new Paragraph({
        children: [new TextRun(b.text)],
        bullet: { level: 0 },
      });
    case "paragraph":
    default:
      return new Paragraph({ children: [new TextRun(b.text)] });
  }
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
