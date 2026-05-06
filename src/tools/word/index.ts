import { z } from "zod";
import type { ToolDefinition, ToolContext } from "../../types.js";
import { DrivePath, Filename } from "../../util/schema.js";
import {
  appendBullets,
  appendHeading,
  appendParagraph,
  applyTemplateChanges,
  createFromBlocks,
  deleteParagraph,
  type DocxBlock,
  extractText,
  insertParagraphAt,
  listParagraphs,
  replaceText,
} from "../../ooxml/docx.js";

/**
 * Word tools download the .docx, manipulate document.xml (OOXML) via the
 * helpers in ../../ooxml/docx.ts, then upload the result.
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

function ensureDocxExt(name: string): string {
  return name.endsWith(".docx") ? name : `${name}.docx`;
}

const BlockSpec: z.ZodType<DocxBlock> = z.object({
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
      return { text: extractText(buf) };
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
      return { paragraphs: listParagraphs(buf) };
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
      const before = await downloadDocx(ctx, input);
      const { buf, count } = replaceText(before, input.replacements);
      await uploadDocx(ctx, input, buf);
      return { replacementsApplied: count };
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
      const before = await downloadDocx(ctx, input);
      const after = appendParagraph(before, input.text);
      await uploadDocx(ctx, input, after);
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
      parentPath: DrivePath.describe("Parent folder path, e.g. '/Reports'"),
      filename: Filename.describe("Filename including .docx extension"),
      blocks: z.array(BlockSpec).min(1),
    }),
    handler: async (input, ctx) => {
      const buf = await createFromBlocks(input.blocks);
      const result = await uploadNewDocxToPath(
        ctx,
        {
          driveId: input.driveId,
          siteId: input.siteId,
          parentPath: input.parentPath,
          filename: ensureDocxExt(input.filename),
        },
        buf,
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
        templatePath: DrivePath.optional().describe("Template path, e.g. '/Templates/proposal.docx'"),
        driveId: z.string().optional(),
        siteId: z.string().optional(),
        parentPath: DrivePath.describe("Destination folder path"),
        filename: Filename.describe("New filename, including .docx"),
        replacements: z
          .array(z.object({ find: z.string().min(1), replace: z.string() }))
          .optional(),
        appendBlocks: z.array(BlockSpec).optional(),
      })
      .refine((d) => Boolean(d.templateItemId) || Boolean(d.templatePath), {
        message: "Provide templateItemId or templatePath",
      }),
    handler: async (input, ctx) => {
      const before = await downloadTemplateBytes(ctx, input);
      const after = applyTemplateChanges(before, {
        replacements: input.replacements,
        appendBlocks: input.appendBlocks,
      });
      const result = await uploadNewDocxToPath(
        ctx,
        {
          driveId: input.driveId,
          siteId: input.siteId,
          parentPath: input.parentPath,
          filename: ensureDocxExt(input.filename),
        },
        after,
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
      const before = await downloadDocx(ctx, input);
      const after = appendHeading(before, input.text, input.level);
      await uploadDocx(ctx, input, after);
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
      const before = await downloadDocx(ctx, input);
      const after = appendBullets(before, input.items);
      await uploadDocx(ctx, input, after);
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
      const before = await downloadDocx(ctx, input);
      const after = insertParagraphAt(before, input.text, input.after);
      await uploadDocx(ctx, input, after);
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
      const before = await downloadDocx(ctx, input);
      const after = deleteParagraph(before, input.paragraphIndex);
      await uploadDocx(ctx, input, after);
      return { ok: true, deletedIndex: input.paragraphIndex };
    },
  },
];
