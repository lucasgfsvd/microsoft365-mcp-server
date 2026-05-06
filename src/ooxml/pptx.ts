import PizZip from "pizzip";
import { decodeXmlEntities, escapeRegex, escapeXml } from "./escape.js";

/** A title-and-bullets slide spec (notes are added separately when creating
 *  brand-new decks via pptxgenjs). */
export type SlideSpec = {
  title?: string;
  bullets?: string[];
};

/** Sorted slide XML file paths inside a .pptx zip, ordered by slide number. */
export function listSlideFiles(zip: PizZip): string[] {
  return Object.keys((zip as unknown as { files: Record<string, unknown> }).files)
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort(
      (a, b) =>
        parseInt(a.match(/slide(\d+)\.xml$/)![1]!) -
        parseInt(b.match(/slide(\d+)\.xml$/)![1]!),
    );
}

/** Concatenate every `<a:t>` run in a slide's XML, separated by spaces. */
export function extractSlideText(xml: string): string {
  return [...xml.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g)]
    .map((m) => decodeXmlEntities(m[1] ?? ""))
    .join(" ");
}

/** All slides, in order. */
export function extractAllSlides(buf: Buffer): { index: number; name: string; text: string }[] {
  const zip = new PizZip(buf);
  return listSlideFiles(zip).map((name, i) => ({
    index: i + 1,
    name,
    text: extractSlideText(zip.file(name)!.asText()),
  }));
}

/** Plain text of a single slide (1-based). Throws if out of range. */
export function getSlideText(buf: Buffer, slideIndex: number): string {
  const zip = new PizZip(buf);
  const names = listSlideFiles(zip);
  const name = names[slideIndex - 1];
  if (!name) throw new Error(`Slide ${slideIndex} not found (deck has ${names.length}).`);
  return extractSlideText(zip.file(name)!.asText());
}

/** Find-and-replace across the deck's `<a:t>` runs. */
export function replaceText(
  buf: Buffer,
  replacements: ReadonlyArray<{ find: string; replace: string }>,
  slideIndexes?: ReadonlyArray<number>,
): { buf: Buffer; replacementsApplied: number; slidesProcessed: number } {
  const zip = new PizZip(buf);
  const names = listSlideFiles(zip);
  const targets = slideIndexes
    ? slideIndexes.map((i) => names[i - 1]).filter((n): n is string => Boolean(n))
    : names;
  let total = 0;
  for (const name of targets) {
    let xml = zip.file(name)!.asText();
    for (const { find, replace } of replacements) {
      xml = xml.replace(/<a:t[^>]*>([^<]*)<\/a:t>/g, (full, inner: string) => {
        if (!inner.includes(find)) return full;
        total += (inner.match(new RegExp(escapeRegex(find), "g")) ?? []).length;
        return full.replace(inner, inner.split(find).join(escapeXml(replace)));
      });
    }
    zip.file(name, xml);
  }
  const out = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
  return { buf: out, replacementsApplied: total, slidesProcessed: targets.length };
}

/** Append a new slide built from a minimal layout (title + bullets, no master). */
export function appendSlide(
  buf: Buffer,
  spec: SlideSpec,
): { buf: Buffer; newIndex: number } {
  const zip = new PizZip(buf);
  const newIndex = appendMinimalSlideToZip(zip, spec.title ?? "", spec.bullets ?? []);
  const out = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
  return { buf: out, newIndex };
}

/** Delete the Nth slide and clean up its references. */
export function deleteSlide(buf: Buffer, slideIndex: number): Buffer {
  const zip = new PizZip(buf);
  const names = listSlideFiles(zip);
  const target = names[slideIndex - 1];
  if (!target) throw new Error(`Slide ${slideIndex} not found (deck has ${names.length}).`);
  const slideNum = parseInt(target.match(/slide(\d+)\.xml$/)![1]!);

  zip.remove(target);
  zip.remove(`ppt/slides/_rels/slide${slideNum}.xml.rels`);

  const presRelsName = "ppt/_rels/presentation.xml.rels";
  const presRels = zip.file(presRelsName);
  if (!presRels) throw new Error("ppt/_rels/presentation.xml.rels missing");
  let presRelsXml = presRels.asText();
  const relRe = new RegExp(
    `<Relationship[^>]*Target="slides/slide${slideNum}\\.xml"[^>]*/>`,
  );
  const relMatch = presRelsXml.match(relRe);
  let removedRId: string | null = null;
  if (relMatch) {
    const idMatch = relMatch[0].match(/Id="(rId\d+)"/);
    removedRId = idMatch ? idMatch[1]! : null;
    presRelsXml = presRelsXml.replace(relRe, "");
    zip.file(presRelsName, presRelsXml);
  }

  const presName = "ppt/presentation.xml";
  const presFile = zip.file(presName);
  if (presFile && removedRId) {
    let presXml = presFile.asText();
    presXml = presXml.replace(new RegExp(`<p:sldId[^>]*r:id="${removedRId}"[^>]*/>`), "");
    zip.file(presName, presXml);
  }

  const ctName = "[Content_Types].xml";
  const ct = zip.file(ctName);
  if (ct) {
    const ctXml = ct
      .asText()
      .replace(
        new RegExp(`<Override[^>]*PartName="/ppt/slides/slide${slideNum}\\.xml"[^>]*/>`),
        "",
      );
    zip.file(ctName, ctXml);
  }

  return zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
}

/** Apply template-style changes (replacements + appended slides). */
export function applyTemplateChanges(
  buf: Buffer,
  opts: {
    replacements?: ReadonlyArray<{ find: string; replace: string }>;
    appendSlides?: ReadonlyArray<SlideSpec>;
  },
): Buffer {
  const zip = new PizZip(buf);

  if (opts.replacements?.length) {
    for (const name of listSlideFiles(zip)) {
      let xml = zip.file(name)!.asText();
      for (const { find, replace } of opts.replacements) {
        xml = xml.replace(/<a:t[^>]*>([^<]*)<\/a:t>/g, (full, inner: string) => {
          if (!inner.includes(find)) return full;
          return full.replace(inner, inner.split(find).join(escapeXml(replace)));
        });
      }
      zip.file(name, xml);
    }
  }

  if (opts.appendSlides?.length) {
    for (const slide of opts.appendSlides) {
      appendMinimalSlideToZip(zip, slide.title ?? "", slide.bullets ?? []);
    }
  }

  return zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
}

function appendMinimalSlideToZip(zip: PizZip, title: string, bullets: ReadonlyArray<string>): number {
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

function buildMinimalSlideXml(title: string, bullets: ReadonlyArray<string>): string {
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
  const bodyRun =
    `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
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
