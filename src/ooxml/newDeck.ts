import PptxGenJS from "pptxgenjs";
import PizZip from "pizzip";
import type { SlideSpec } from "./pptx.js";

export interface NewSlide extends SlideSpec {
  notes?: string;
}

const MASTER = "TITLE_AND_BODY";

/**
 * Build a brand-new 16:9 deck.
 *
 * Titles and bullets go into real title and body placeholders from a slide
 * master, not free text boxes: that is what makes PowerPoint's outline view,
 * its accessibility checker and screen readers see each slide's title.
 */
export async function buildDeck(slides: NewSlide[], meta: { title?: string; author?: string } = {}): Promise<Buffer> {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_WIDE";
  if (meta.title) pres.title = meta.title;
  if (meta.author) pres.author = meta.author;
  pres.defineSlideMaster({
    title: MASTER,
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 12.3, h: 1, fontSize: 32, bold: true }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.5, y: 1.5, w: 12.3, h: 5.5, fontSize: 18, valign: "top" }, text: "" } },
    ],
  });

  for (const spec of slides) {
    const slide = pres.addSlide({ masterName: MASTER });
    if (spec.title) slide.addText(spec.title, { placeholder: "title" });
    if (spec.bullets?.length) {
      slide.addText(
        spec.bullets.map((t) => ({ text: t, options: { bullet: true } })),
        { placeholder: "body" },
      );
    }
    if (spec.notes) slide.addNotes(spec.notes);
  }

  const out = (await pres.write({ outputType: "nodebuffer" })) as Buffer;
  return normalizeTitlePlaceholders(Buffer.isBuffer(out) ? out : Buffer.from(out));
}

/**
 * pptxgenjs gives title placeholders an explicit idx (100). PowerPoint writes a
 * title with no idx (so 0), and tools that follow its convention — python-pptx
 * among them — find the title by idx 0 and report these slides as untitled.
 * Dropping the idx on both the slides and their layout keeps the two linked.
 */
export function normalizeTitlePlaceholders(buf: Buffer): Buffer {
  const zip = new PizZip(buf);
  const parts = Object.keys((zip as unknown as { files: Record<string, unknown> }).files).filter((n) =>
    /^ppt\/(slides\/slide|slideLayouts\/slideLayout)\d+\.xml$/.test(n),
  );
  for (const name of parts) {
    const xml = zip.file(name)!.asText();
    const fixed = xml.replace(/<p:ph\s+idx="\d+"\s+type="title"/g, '<p:ph type="title"');
    if (fixed !== xml) zip.file(name, fixed);
  }
  return zip.generate({ type: "nodebuffer" });
}
