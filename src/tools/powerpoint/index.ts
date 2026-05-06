import { z } from "zod";
import PptxGenJS from "pptxgenjs";
import type { ToolDefinition, ToolContext } from "../../types.js";
import { DrivePath, Filename } from "../../util/schema.js";
import {
  appendSlide,
  applyTemplateChanges,
  deleteSlide,
  extractAllSlides,
  getSlideText,
  replaceText,
  type SlideSpec as MinimalSlideSpec,
} from "../../ooxml/pptx.js";

const DeckRef = z.object({
  driveId: z.string().optional(),
  siteId: z.string().optional(),
  itemId: z.string().describe("driveItem id of the .pptx file"),
});

function drivePrefix(r: { driveId?: string; siteId?: string }): string {
  if (r.driveId) return `/drives/${r.driveId}`;
  if (r.siteId) return `/sites/${r.siteId}/drive`;
  return `/me/drive`;
}

async function downloadPptx(ctx: ToolContext, ref: z.infer<typeof DeckRef>): Promise<Buffer> {
  const base = drivePrefix(ref);
  const stream: NodeJS.ReadableStream = await ctx.graph.api(`${base}/items/${ref.itemId}/content`).getStream();
  const chunks: Buffer[] = [];
  for await (const c of stream) chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c));
  return Buffer.concat(chunks);
}

async function uploadPptx(ctx: ToolContext, ref: z.infer<typeof DeckRef>, buf: Buffer): Promise<unknown> {
  const base = drivePrefix(ref);
  return ctx.graph.api(`${base}/items/${ref.itemId}/content`).put(buf);
}

async function uploadNewPptxToPath(
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

function ensurePptxExt(name: string): string {
  return name.endsWith(".pptx") ? name : `${name}.pptx`;
}

const SlideSpec = z.object({
  title: z.string().optional(),
  bullets: z.array(z.string()).optional(),
  notes: z.string().optional(),
});

export const powerpointTools: ToolDefinition[] = [
  {
    name: "powerpoint_list_slides",
    surface: "powerpoint",
    description: "List slides in a .pptx with extracted text content per slide.",
    requiredScopes: ["Files.Read.All", "Sites.Read.All"],
    inputSchema: DeckRef,
    handler: async (input, ctx) => {
      const buf = await downloadPptx(ctx, input);
      return { slides: extractAllSlides(buf) };
    },
  },
  {
    name: "powerpoint_get_slide_text",
    surface: "powerpoint",
    description: "Get text of a single slide by 1-based index.",
    requiredScopes: ["Files.Read.All", "Sites.Read.All"],
    inputSchema: DeckRef.extend({ slideIndex: z.number().int().min(1) }),
    handler: async (input, ctx) => {
      const buf = await downloadPptx(ctx, input);
      return { index: input.slideIndex, text: getSlideText(buf, input.slideIndex) };
    },
  },
  {
    name: "powerpoint_extract_all_text",
    surface: "powerpoint",
    description: "Return all text in the deck joined with slide separators (useful for summarization).",
    requiredScopes: ["Files.Read.All", "Sites.Read.All"],
    inputSchema: DeckRef,
    handler: async (input, ctx) => {
      const buf = await downloadPptx(ctx, input);
      const text = extractAllSlides(buf)
        .map((s) => `--- Slide ${s.index} ---\n${s.text}`)
        .join("\n\n");
      return { text };
    },
  },
  {
    name: "powerpoint_replace_text",
    surface: "powerpoint",
    description:
      "Find-and-replace text across all slides. Matches against each <a:t> run (may miss matches split across runs).",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: DeckRef.extend({
      replacements: z.array(z.object({ find: z.string().min(1), replace: z.string() })).min(1),
      slideIndexes: z.array(z.number().int().min(1)).optional().describe("Limit to specific slides; omit for all."),
    }),
    handler: async (input, ctx) => {
      const before = await downloadPptx(ctx, input);
      const { buf, replacementsApplied, slidesProcessed } = replaceText(
        before,
        input.replacements,
        input.slideIndexes,
      );
      await uploadPptx(ctx, input, buf);
      return { replacementsApplied, slidesProcessed };
    },
  },
  {
    name: "powerpoint_create_deck",
    surface: "powerpoint",
    description:
      "Create a brand new .pptx from a list of slide specs and upload it to OneDrive/SharePoint. " +
      "Each slide supports a title, bullet list, and speaker notes. Uses pptxgenjs (16:9, default theme). " +
      "Returns the driveItem metadata of the new file.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: z.object({
      driveId: z.string().optional(),
      siteId: z.string().optional(),
      parentPath: DrivePath.describe("Parent folder path, e.g. '/Decks' or '/'"),
      filename: Filename.describe("Filename including .pptx extension, e.g. 'Q3-review.pptx'"),
      title: z.string().optional().describe("Deck title (stored in .pptx metadata)."),
      author: z.string().optional(),
      slides: z.array(SlideSpec).min(1),
    }),
    handler: async (input, ctx) => {
      const pres = new PptxGenJS();
      pres.layout = "LAYOUT_WIDE";
      if (input.title) pres.title = input.title;
      if (input.author) pres.author = input.author;
      for (const spec of input.slides) {
        const slide = pres.addSlide();
        if (spec.title) {
          slide.addText(spec.title, { x: 0.5, y: 0.3, w: 12, h: 1, fontSize: 32, bold: true });
        }
        if (spec.bullets && spec.bullets.length) {
          slide.addText(
            (spec.bullets as string[]).map((t: string) => ({ text: t, options: { bullet: true } })),
            { x: 0.5, y: 1.5, w: 12, h: 5.5, fontSize: 18, valign: "top" },
          );
        }
        if (spec.notes) slide.addNotes(spec.notes);
      }
      const out = (await pres.write({ outputType: "nodebuffer" })) as Buffer;
      const result = await uploadNewPptxToPath(
        ctx,
        {
          driveId: input.driveId,
          siteId: input.siteId,
          parentPath: input.parentPath,
          filename: ensurePptxExt(input.filename),
        },
        Buffer.isBuffer(out) ? out : Buffer.from(out),
      );
      return { ok: true, slides: input.slides.length, driveItem: result };
    },
  },
  {
    name: "powerpoint_add_slide",
    surface: "powerpoint",
    description:
      "Append a new slide to an existing .pptx with a title and optional bullets. The new slide is built from " +
      "a minimal layout (title + body) independent of the deck's master — visual style may differ from " +
      "existing slides. Use powerpoint_create_deck + powerpoint_replace_text for richer templating.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: DeckRef.extend({
      title: z.string().optional(),
      bullets: z.array(z.string()).optional(),
      notes: z.string().optional(),
    }),
    handler: async (input, ctx) => {
      const before = await downloadPptx(ctx, input);
      const spec: MinimalSlideSpec = { title: input.title, bullets: input.bullets };
      const { buf, newIndex } = appendSlide(before, spec);
      await uploadPptx(ctx, input, buf);
      return { ok: true, newSlideIndex: newIndex };
    },
  },
  {
    name: "powerpoint_create_from_template",
    surface: "powerpoint",
    description:
      "Create a new .pptx by copying a template deck, applying placeholder replacements (e.g. {{CLIENT}} → 'Acme'), " +
      "and optionally appending extra slides. Theme, layouts, masters, and existing slide formatting from the " +
      "template are preserved. Identify the template by templateItemId OR templatePath.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: z
      .object({
        templateDriveId: z.string().optional(),
        templateSiteId: z.string().optional(),
        templateItemId: z.string().optional(),
        templatePath: DrivePath.optional().describe("Template path, e.g. '/Templates/proposal.pptx'"),
        driveId: z.string().optional().describe("Destination drive (defaults to the user's OneDrive)."),
        siteId: z.string().optional(),
        parentPath: DrivePath.describe("Destination folder path"),
        filename: Filename.describe("New filename, including .pptx"),
        replacements: z
          .array(z.object({ find: z.string().min(1), replace: z.string() }))
          .optional()
          .describe("Applied across every slide's <a:t> runs."),
        appendSlides: z
          .array(SlideSpec)
          .optional()
          .describe("Slides appended after the template's existing slides (minimal layout)."),
      })
      .refine((d) => Boolean(d.templateItemId) || Boolean(d.templatePath), {
        message: "Provide templateItemId or templatePath",
      }),
    handler: async (input, ctx) => {
      const before = await downloadTemplateBytes(ctx, input);
      const after = applyTemplateChanges(before, {
        replacements: input.replacements,
        appendSlides: input.appendSlides?.map(
          (s: z.infer<typeof SlideSpec>): MinimalSlideSpec => ({ title: s.title, bullets: s.bullets }),
        ),
      });
      const result = await uploadNewPptxToPath(
        ctx,
        {
          driveId: input.driveId,
          siteId: input.siteId,
          parentPath: input.parentPath,
          filename: ensurePptxExt(input.filename),
        },
        after,
      );
      return { ok: true, driveItem: result };
    },
  },
  {
    name: "powerpoint_delete_slide",
    surface: "powerpoint",
    description: "Delete a slide by 1-based index from an existing .pptx.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: DeckRef.extend({ slideIndex: z.number().int().min(1) }),
    handler: async (input, ctx) => {
      const before = await downloadPptx(ctx, input);
      const after = deleteSlide(before, input.slideIndex);
      await uploadPptx(ctx, input, after);
      return { ok: true, deletedSlideIndex: input.slideIndex };
    },
  },
];
