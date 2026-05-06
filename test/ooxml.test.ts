import { describe, it, expect } from "vitest";
import PptxGenJS from "pptxgenjs";
import PizZip from "pizzip";
import { escapeRegex, escapeXml } from "../src/ooxml/escape.js";
import {
  appendBullets,
  appendHeading,
  appendParagraph,
  applyTemplateChanges as applyDocxTemplate,
  createFromBlocks,
  deleteParagraph,
  extractText,
  insertParagraphAt,
  listParagraphs,
  replaceText as replaceDocxText,
} from "../src/ooxml/docx.js";
import {
  appendSlide,
  applyTemplateChanges as applyPptxTemplate,
  deleteSlide,
  extractAllSlides,
  getSlideText,
  listSlideFiles,
  replaceText as replacePptxText,
} from "../src/ooxml/pptx.js";

// ---------------------------------------------------------------------------
// Fixtures — generated at test time, not committed.
// ---------------------------------------------------------------------------

function makeDocx(): Promise<Buffer> {
  return createFromBlocks([
    { kind: "title", text: "Test Title" },
    { kind: "paragraph", text: "First paragraph" },
    { kind: "paragraph", text: "Second paragraph with {{PLACEHOLDER}}" },
  ]);
}

async function makePptx(): Promise<Buffer> {
  const pres = new PptxGenJS();
  pres.addSlide().addText("Slide One", { x: 0.5, y: 0.5 });
  const second = pres.addSlide();
  second.addText("Slide Two", { x: 0.5, y: 0.5 });
  second.addText("Has {{PLACEHOLDER}} text", { x: 0.5, y: 1.5 });
  return (await pres.write({ outputType: "nodebuffer" })) as Buffer;
}

// ---------------------------------------------------------------------------
// escape
// ---------------------------------------------------------------------------

describe("escape", () => {
  it("escapeXml covers all five XML metacharacters", () => {
    expect(escapeXml("&<>\"'")).toBe("&amp;&lt;&gt;&quot;&apos;");
  });

  it("escapeXml is safe to splice into an attribute value", () => {
    const evil = `" onclick="alert(1)`;
    expect(escapeXml(evil)).not.toContain('"');
    expect(escapeXml(evil)).toContain("&quot;");
  });

  it("escapeRegex escapes regex special chars", () => {
    expect(escapeRegex("a.b*c+?^${}()|[]\\")).toBe(
      "a\\.b\\*c\\+\\?\\^\\$\\{\\}\\(\\)\\|\\[\\]\\\\",
    );
  });
});

// ---------------------------------------------------------------------------
// docx
// ---------------------------------------------------------------------------

describe("ooxml/docx", () => {
  it("extractText round-trips paragraphs from createFromBlocks", async () => {
    const buf = await makeDocx();
    const text = extractText(buf);
    expect(text).toContain("Test Title");
    expect(text).toContain("First paragraph");
    expect(text).toContain("Second paragraph with {{PLACEHOLDER}}");
  });

  it("listParagraphs splits on paragraph boundaries", async () => {
    const buf = await makeDocx();
    const paras = listParagraphs(buf);
    expect(paras.some((p) => p.includes("Test Title"))).toBe(true);
    expect(paras.some((p) => p.includes("First paragraph"))).toBe(true);
  });

  it("replaceText substitutes literal occurrences and reports count", async () => {
    const buf = await makeDocx();
    const { buf: out, count } = replaceDocxText(buf, [{ find: "{{PLACEHOLDER}}", replace: "Acme Corp" }]);
    expect(count).toBe(1);
    expect(extractText(out)).toContain("Acme Corp");
    expect(extractText(out)).not.toContain("{{PLACEHOLDER}}");
  });

  it("replaceText escapes XML metacharacters in replacement text", async () => {
    const buf = await makeDocx();
    const { buf: out } = replaceDocxText(buf, [
      { find: "{{PLACEHOLDER}}", replace: `<script>"&'</script>` },
    ]);
    // The visible (extracted) text should be the literal injection;
    // the underlying XML should contain the escaped form.
    expect(extractText(out)).toContain(`<script>"&'</script>`);
    const zip = new PizZip(out);
    const xml = zip.file("word/document.xml")!.asText();
    expect(xml).not.toContain("<script>");
    expect(xml).toContain("&lt;script&gt;");
    expect(xml).toContain("&quot;");
    expect(xml).toContain("&apos;");
  });

  it("replaceText handles regex-meta find strings without crashing", async () => {
    const buf = await makeDocx();
    const { count } = replaceDocxText(buf, [{ find: "{{PLACEHOLDER}}", replace: "ok" }]);
    // Find string contains regex metas ({}). Should still match literally.
    expect(count).toBe(1);
  });

  it("appendParagraph adds at the end", async () => {
    const buf = await makeDocx();
    const out = appendParagraph(buf, "Appended line");
    const paras = listParagraphs(out);
    expect(paras[paras.length - 1]).toBe("Appended line");
  });

  it("appendHeading uses the requested style", async () => {
    const buf = await makeDocx();
    const out = appendHeading(buf, "My Section", 2);
    const xml = new PizZip(out).file("word/document.xml")!.asText();
    expect(xml).toContain('w:val="Heading2"');
    expect(xml).toContain("My Section");
  });

  it("appendBullets adds list-styled paragraphs", async () => {
    const buf = await makeDocx();
    const out = appendBullets(buf, ["one", "two", "three"]);
    const xml = new PizZip(out).file("word/document.xml")!.asText();
    expect(xml).toContain('w:val="ListParagraph"');
    expect((xml.match(/<w:numPr>/g) ?? []).length).toBeGreaterThanOrEqual(3);
    const text = extractText(out);
    expect(text).toContain("one");
    expect(text).toContain("two");
    expect(text).toContain("three");
  });

  it("insertParagraphAt(after=0) inserts at the top", async () => {
    const buf = await makeDocx();
    const out = insertParagraphAt(buf, "TOP", 0);
    const paras = listParagraphs(out);
    expect(paras[0]).toBe("TOP");
  });

  it("insertParagraphAt(after=N) inserts after the Nth paragraph", async () => {
    const buf = await makeDocx();
    const before = listParagraphs(buf);
    const out = insertParagraphAt(buf, "INSERTED", 2);
    const after = listParagraphs(out);
    expect(after.length).toBe(before.length + 1);
    expect(after[2]).toBe("INSERTED");
  });

  it("deleteParagraph removes the Nth paragraph (1-based)", async () => {
    const buf = await makeDocx();
    const before = listParagraphs(buf);
    const out = deleteParagraph(buf, 1);
    const after = listParagraphs(out);
    expect(after.length).toBe(before.length - 1);
    expect(after.every((p) => !p.includes("Test Title"))).toBe(true);
  });

  it("applyTemplateChanges combines replacements and appended blocks", async () => {
    const buf = await makeDocx();
    const out = applyDocxTemplate(buf, {
      replacements: [{ find: "{{PLACEHOLDER}}", replace: "Globex" }],
      appendBlocks: [
        { kind: "heading2", text: "Appendix" },
        { kind: "bullet", text: "extra item" },
      ],
    });
    const text = extractText(out);
    expect(text).toContain("Globex");
    expect(text).not.toContain("{{PLACEHOLDER}}");
    expect(text).toContain("Appendix");
    expect(text).toContain("extra item");
  });
});

// ---------------------------------------------------------------------------
// pptx
// ---------------------------------------------------------------------------

describe("ooxml/pptx", () => {
  it("listSlideFiles enumerates slides in order", async () => {
    const buf = await makePptx();
    const zip = new PizZip(buf);
    const names = listSlideFiles(zip);
    expect(names.length).toBe(2);
    expect(names[0]).toMatch(/slide1\.xml$/);
    expect(names[1]).toMatch(/slide2\.xml$/);
  });

  it("extractAllSlides returns indexed text per slide", async () => {
    const buf = await makePptx();
    const slides = extractAllSlides(buf);
    expect(slides.length).toBe(2);
    expect(slides[0]?.text).toContain("Slide One");
    expect(slides[1]?.text).toContain("Slide Two");
  });

  it("getSlideText returns text for a valid index", async () => {
    const buf = await makePptx();
    expect(getSlideText(buf, 1)).toContain("Slide One");
  });

  it("getSlideText throws for an out-of-range index", async () => {
    const buf = await makePptx();
    expect(() => getSlideText(buf, 99)).toThrow(/Slide 99 not found/);
  });

  it("replaceText replaces across slides and reports counts", async () => {
    const buf = await makePptx();
    const { buf: out, replacementsApplied, slidesProcessed } = replacePptxText(buf, [
      { find: "{{PLACEHOLDER}}", replace: "Acme" },
    ]);
    expect(replacementsApplied).toBe(1);
    expect(slidesProcessed).toBe(2);
    const text = extractAllSlides(out)
      .map((s) => s.text)
      .join(" ");
    expect(text).toContain("Acme");
    expect(text).not.toContain("{{PLACEHOLDER}}");
  });

  it("replaceText respects slideIndexes filter", async () => {
    const buf = await makePptx();
    const { replacementsApplied, slidesProcessed } = replacePptxText(
      buf,
      [{ find: "{{PLACEHOLDER}}", replace: "Acme" }],
      [1], // only slide 1, which doesn't have the placeholder
    );
    expect(replacementsApplied).toBe(0);
    expect(slidesProcessed).toBe(1);
  });

  it("appendSlide adds a slide and returns its 1-based index", async () => {
    const buf = await makePptx();
    const { buf: out, newIndex } = appendSlide(buf, {
      title: "New Slide",
      bullets: ["alpha", "beta"],
    });
    expect(newIndex).toBe(3);
    const slides = extractAllSlides(out);
    expect(slides.length).toBe(3);
    expect(slides[2]?.text).toContain("New Slide");
    expect(slides[2]?.text).toContain("alpha");
  });

  it("appendSlide registers the slide in [Content_Types].xml and presentation.xml.rels", async () => {
    const buf = await makePptx();
    const { buf: out } = appendSlide(buf, { title: "X", bullets: [] });
    const zip = new PizZip(out);
    expect(zip.file("[Content_Types].xml")!.asText()).toContain("slide3.xml");
    expect(zip.file("ppt/_rels/presentation.xml.rels")!.asText()).toContain("slides/slide3.xml");
  });

  it("deleteSlide removes the slide and its references", async () => {
    const buf = await makePptx();
    const out = deleteSlide(buf, 1);
    const slides = extractAllSlides(out);
    expect(slides.length).toBe(1);
    expect(slides[0]?.text).toContain("Slide Two");

    const zip = new PizZip(out);
    expect(zip.file("ppt/slides/slide1.xml")).toBeNull();
    const presRels = zip.file("ppt/_rels/presentation.xml.rels")!.asText();
    expect(presRels).not.toMatch(/Target="slides\/slide1\.xml"/);
  });

  it("deleteSlide throws for an out-of-range index", async () => {
    const buf = await makePptx();
    expect(() => deleteSlide(buf, 99)).toThrow(/Slide 99 not found/);
  });

  it("applyTemplateChanges runs replacements then appends slides", async () => {
    const buf = await makePptx();
    const out = applyPptxTemplate(buf, {
      replacements: [{ find: "{{PLACEHOLDER}}", replace: "Globex" }],
      appendSlides: [{ title: "Thanks", bullets: ["q&a"] }],
    });
    const slides = extractAllSlides(out);
    expect(slides.length).toBe(3);
    expect(slides.map((s) => s.text).join(" ")).toContain("Globex");
    expect(slides[2]?.text).toContain("Thanks");
  });
});
