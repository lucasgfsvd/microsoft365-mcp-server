import { z } from "zod";
import PizZip from "pizzip";
import PptxGenJS from "pptxgenjs";
import type { ToolDefinition, ToolContext } from "../../types.js";

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

function listSlideFiles(zip: PizZip): string[] {
  return Object.keys((zip as unknown as { files: Record<string, unknown> }).files)
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => parseInt(a.match(/slide(\d+)\.xml$/)![1]!) - parseInt(b.match(/slide(\d+)\.xml$/)![1]!));
}

function extractSlideText(xml: string): string {
  return [...xml.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g)].map((m) => m[1]).join(" ");
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
      const zip = new PizZip(buf);
      const slides = listSlideFiles(zip).map((name, i) => ({
        index: i + 1,
        name,
        text: extractSlideText(zip.file(name)!.asText()),
      }));
      return { slides };
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
      const zip = new PizZip(buf);
      const names = listSlideFiles(zip);
      const name = names[input.slideIndex - 1];
      if (!name) throw new Error(`Slide ${input.slideIndex} not found (deck has ${names.length}).`);
      return { index: input.slideIndex, text: extractSlideText(zip.file(name)!.asText()) };
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
      const zip = new PizZip(buf);
      const names = listSlideFiles(zip);
      const text = names
        .map((n, i) => `--- Slide ${i + 1} ---\n${extractSlideText(zip.file(n)!.asText())}`)
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
      const buf = await downloadPptx(ctx, input);
      const zip = new PizZip(buf);
      const names = listSlideFiles(zip);
      const targets = input.slideIndexes
        ? input.slideIndexes.map((i: number) => names[i - 1]).filter((n: string | undefined): n is string => Boolean(n))
        : names;
      let total = 0;
      for (const name of targets) {
        let xml = zip.file(name)!.asText();
        for (const { find, replace } of input.replacements) {
          xml = xml.replace(/<a:t[^>]*>([^<]*)<\/a:t>/g, (full, inner: string) => {
            if (!inner.includes(find)) return full;
            total += (inner.match(new RegExp(escapeRegex(find), "g")) ?? []).length;
            return full.replace(inner, inner.split(find).join(escapeXml(replace)));
          });
        }
        zip.file(name, xml);
      }
      const out = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
      await uploadPptx(ctx, input, out);
      return { replacementsApplied: total, slidesProcessed: targets.length };
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
      parentPath: z.string().describe("Parent folder path, e.g. '/Decks' or '/'"),
      filename: z.string().describe("Filename including .pptx extension, e.g. 'Q3-review.pptx'"),
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
      // pptxgenjs returns Uint8Array/Buffer for nodebuffer output.
      const out = (await pres.write({ outputType: "nodebuffer" })) as Buffer;
      const filename = input.filename.endsWith(".pptx") ? input.filename : `${input.filename}.pptx`;
      const result = await uploadNewPptxToPath(
        ctx,
        { driveId: input.driveId, siteId: input.siteId, parentPath: input.parentPath, filename },
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
      const buf = await downloadPptx(ctx, input);
      const zip = new PizZip(buf);
      const newIndex = appendMinimalSlide(zip, input.title ?? "", input.bullets ?? []);
      const out = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
      await uploadPptx(ctx, input, out);
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
        templatePath: z.string().optional().describe("Template path, e.g. '/Templates/proposal.pptx'"),
        driveId: z.string().optional().describe("Destination drive (defaults to the user's OneDrive)."),
        siteId: z.string().optional(),
        parentPath: z.string().describe("Destination folder path"),
        filename: z.string().describe("New filename, including .pptx"),
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
      const buf = await downloadTemplateBytes(ctx, input);
      const zip = new PizZip(buf);

      if (input.replacements && input.replacements.length) {
        const names = listSlideFiles(zip);
        for (const name of names) {
          let xml = zip.file(name)!.asText();
          for (const { find, replace } of input.replacements) {
            xml = xml.replace(/<a:t[^>]*>([^<]*)<\/a:t>/g, (full, inner: string) => {
              if (!inner.includes(find)) return full;
              return full.replace(inner, inner.split(find).join(escapeXml(replace)));
            });
          }
          zip.file(name, xml);
        }
      }

      if (input.appendSlides && input.appendSlides.length) {
        for (const spec of input.appendSlides as z.infer<typeof SlideSpec>[]) {
          appendMinimalSlide(zip, spec.title ?? "", spec.bullets ?? []);
        }
      }

      const out = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
      const filename = input.filename.endsWith(".pptx") ? input.filename : `${input.filename}.pptx`;
      const result = await uploadNewPptxToPath(
        ctx,
        { driveId: input.driveId, siteId: input.siteId, parentPath: input.parentPath, filename },
        out,
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
      const buf = await downloadPptx(ctx, input);
      const zip = new PizZip(buf);
      const names = listSlideFiles(zip);
      const target = names[input.slideIndex - 1];
      if (!target) throw new Error(`Slide ${input.slideIndex} not found (deck has ${names.length}).`);
      const slideNum = parseInt(target.match(/slide(\d+)\.xml$/)![1]!);

      // Remove slide + rels.
      zip.remove(target);
      zip.remove(`ppt/slides/_rels/slide${slideNum}.xml.rels`);

      // Find the rId pointing to this slide in presentation.xml.rels, then remove it.
      const presRelsName = "ppt/_rels/presentation.xml.rels";
      const presRels = zip.file(presRelsName);
      if (!presRels) throw new Error("ppt/_rels/presentation.xml.rels missing");
      let presRelsXml = presRels.asText();
      const relRe = new RegExp(
        `<Relationship[^/]*Target="slides/slide${slideNum}\\.xml"[^/]*/>`,
      );
      const relMatch = presRelsXml.match(relRe);
      let removedRId: string | null = null;
      if (relMatch) {
        const idMatch = relMatch[0].match(/Id="(rId\d+)"/);
        removedRId = idMatch ? idMatch[1]! : null;
        presRelsXml = presRelsXml.replace(relRe, "");
        zip.file(presRelsName, presRelsXml);
      }

      // Remove <p:sldId .../> referencing that rId from presentation.xml.
      const presName = "ppt/presentation.xml";
      const presFile = zip.file(presName);
      if (presFile && removedRId) {
        let presXml = presFile.asText();
        presXml = presXml.replace(new RegExp(`<p:sldId[^/]*r:id="${removedRId}"[^/]*/>`), "");
        zip.file(presName, presXml);
      }

      // Remove the Content_Types override for this slide.
      const ctName = "[Content_Types].xml";
      const ct = zip.file(ctName);
      if (ct) {
        const ctXml = ct
          .asText()
          .replace(new RegExp(`<Override[^/]*PartName="/ppt/slides/slide${slideNum}\\.xml"[^/]*/>`), "");
        zip.file(ctName, ctXml);
      }

      const out = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
      await uploadPptx(ctx, input, out);
      return { ok: true, deletedSlideIndex: input.slideIndex };
    },
  },
];

function appendMinimalSlide(zip: PizZip, title: string, bullets: string[]): number {
  const names = listSlideFiles(zip);
  const newIndex = names.length + 1;
  zip.file(`ppt/slides/slide${newIndex}.xml`, buildMinimalSlideXml(title, bullets));
  zip.file(
    `ppt/slides/_rels/slide${newIndex}.xml.rels`,
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
      `</Relationships>`,
  );

  const ctName = "[Content_Types].xml";
  const ct = zip.file(ctName);
  if (!ct) throw new Error("[Content_Types].xml missing");
  const override = `<Override PartName="/ppt/slides/slide${newIndex}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
  zip.file(ctName, ct.asText().replace("</Types>", `${override}</Types>`));

  const presRelsName = "ppt/_rels/presentation.xml.rels";
  const presRels = zip.file(presRelsName);
  if (!presRels) throw new Error("ppt/_rels/presentation.xml.rels missing");
  let presRelsXml = presRels.asText();
  const rIds = [...presRelsXml.matchAll(/Id="rId(\d+)"/g)].map((m) => parseInt(m[1]!));
  const newRId = (rIds.length ? Math.max(...rIds) : 0) + 1;
  const newRel = `<Relationship Id="rId${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${newIndex}.xml"/>`;
  presRelsXml = presRelsXml.replace("</Relationships>", `${newRel}</Relationships>`);
  zip.file(presRelsName, presRelsXml);

  const presName = "ppt/presentation.xml";
  const presFile = zip.file(presName);
  if (!presFile) throw new Error("ppt/presentation.xml missing");
  let presXml = presFile.asText();
  const sldIds = [...presXml.matchAll(/<p:sldId[^>]*\sid="(\d+)"/g)].map((m) => parseInt(m[1]!));
  const newSldId = (sldIds.length ? Math.max(...sldIds) : 255) + 1;
  presXml = presXml.replace(
    "</p:sldIdLst>",
    `<p:sldId id="${newSldId}" r:id="rId${newRId}"/></p:sldIdLst>`,
  );
  zip.file(presName, presXml);

  return newIndex;
}

function buildMinimalSlideXml(title: string, bullets: string[]): string {
  const titleRun = title
    ? `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
      `<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>` +
      `<p:spPr/>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/>` +
      `<a:p><a:r><a:rPr lang="en-US"/><a:t>${escapeXml(title)}</a:t></a:r></a:p>` +
      `</p:txBody></p:sp>`
    : "";
  const bulletParas = bullets.length
    ? bullets
        .map(
          (b) =>
            `<a:p><a:pPr lvl="0"><a:buChar char="•"/></a:pPr>` +
            `<a:r><a:rPr lang="en-US"/><a:t>${escapeXml(b)}</a:t></a:r></a:p>`,
        )
        .join("")
    : `<a:p><a:endParaRPr lang="en-US"/></a:p>`;
  const bodyRun = `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
    `<p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr/>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/>${bulletParas}</p:txBody></p:sp>`;

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
    `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
    `<p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr/>` +
    titleRun +
    bodyRun +
    `</p:spTree></p:cSld>` +
    `</p:sld>`
  );
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
