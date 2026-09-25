import { describe, it, expect } from "vitest";
import PizZip from "pizzip";
import { buildDeck } from "../src/ooxml/newDeck.js";
import { appendSlide, deleteSlide, extractAllSlides } from "../src/ooxml/pptx.js";

const placeholders = (xml: string) => [...xml.matchAll(/<p:ph\b[^>]*>/g)].map((m) => m[0].replace(/\s+/g, " "));
const part = (buf: Buffer, name: string) => new PizZip(buf).file(name)!.asText();
const layoutWithTitle = (buf: Buffer) => {
  const files = Object.keys((new PizZip(buf) as unknown as { files: Record<string, unknown> }).files);
  return files.filter((n) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(n)).map((n) => part(buf, n)).find((x) => x.includes('type="title"'))!;
};

// Regression, found live with python-pptx: slides were free text boxes, so
// outline view and screen readers saw every created slide as untitled.
describe("buildDeck", () => {
  const specs = [
    { title: "Slide One", bullets: ["Point A", "Point B"], notes: "Speaker notes" },
    { title: "Only a title" },
    { bullets: ["No title here"] },
  ];

  it("puts every title in a title placeholder with idx 0, as PowerPoint writes it", async () => {
    const buf = await buildDeck(specs);
    for (const n of [1, 2, 3]) {
      const phs = placeholders(part(buf, `ppt/slides/slide${n}.xml`));
      expect(phs.some((p) => /type="title"/.test(p) && !/idx=/.test(p))).toBe(true);
      expect(phs.some((p) => /type="body"/.test(p))).toBe(true);
    }
    // The layout's title must match the slides', or they stop inheriting from it.
    expect(placeholders(layoutWithTitle(buf)).find((p) => p.includes('type="title"'))).not.toMatch(/idx=/);
  });

  it("places titles and bullets in their placeholders", async () => {
    const buf = await buildDeck(specs);
    const slide1 = part(buf, "ppt/slides/slide1.xml");
    const titleShape = slide1.split("<p:sp>").find((s) => s.includes('type="title"'))!;
    expect(titleShape).toContain("Slide One");
    expect(titleShape).not.toContain("Point A");
    expect(extractAllSlides(buf).map((s) => s.text)).toEqual(["Slide One Point A Point B", "Only a title", "No title here"]);
  });

  it("still works with the in-place slide edits", async () => {
    const buf = await buildDeck(specs);
    const { buf: added, newIndex } = appendSlide(buf, { title: "Added", bullets: ["x"] });
    expect(newIndex).toBe(4);
    const removed = deleteSlide(added, 2);
    expect(extractAllSlides(removed).map((s) => s.text)).toEqual(["Slide One Point A Point B", "No title here", "Added x"]);
  });
});
