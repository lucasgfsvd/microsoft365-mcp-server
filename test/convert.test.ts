import { describe, it, expect } from "vitest";
import { Document, Packer, Paragraph } from "docx";
import { convertForConversation } from "../src/content/convert.js";
import { fileOutput } from "../src/util/toolOutput.js";

/** A one-page PDF saying `text`, built by hand so the test needs no fixture file. */
function pdfSaying(text: string): Buffer {
  const stream = `BT /F1 24 Tf 72 720 Td (${text}) Tj ET`;
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
  ];
  let body = "%PDF-1.4\n";
  const offsets: number[] = [];
  objects.forEach((o, i) => {
    offsets.push(body.length);
    body += `${i + 1} 0 obj\n${o}\nendobj\n`;
  });
  const xref = body.length;
  body += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const off of offsets) body += `${String(off).padStart(10, "0")} 00000 n \n`;
  body += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(body, "latin1");
}

const PNG = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==", "base64");

describe("convertForConversation", () => {
  it("reads a PDF's text, page by page", async () => {
    const c = await convertForConversation(pdfSaying("Quarterly revenue up 12 percent"), "application/pdf", "q3.pdf");
    expect(c).toMatchObject({ kind: "text", mimeType: "text/plain" });
    expect(c.kind === "text" && c.text).toMatch(/^--- Page 1 ---\nQuarterly revenue up 12 percent/);
  });

  it("says when a PDF has no text to give, as a scan has none", async () => {
    const c = await convertForConversation(pdfSaying(""), undefined, "scan.pdf");
    expect(c).toMatchObject({ kind: "text", note: expect.stringMatching(/No text layer/) });
  });

  it("hands back the bytes of a document it cannot open, rather than failing", async () => {
    for (const [mime, name] of [["application/pdf", "bad.pdf"], [undefined, "bad.docx"], [undefined, "bad.xlsx"]] as const) {
      expect(await convertForConversation(Buffer.from("not really"), mime, name), name).toMatchObject({ kind: "binary" });
    }
  });

  it("shows images as images, by type or by name", async () => {
    expect(await convertForConversation(PNG, "image/png", "x")).toEqual({ kind: "image", data: PNG.toString("base64"), mimeType: "image/png" });
    expect(await convertForConversation(PNG, "application/octet-stream", "photo.JPG")).toMatchObject({ kind: "image", mimeType: "image/jpeg" });
    // Not a type models take: left as bytes.
    expect(await convertForConversation(PNG, "image/tiff", "scan.tiff")).toMatchObject({ kind: "binary" });
  });

  it("keeps Word as text and plain text as text", async () => {
    const docx = await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph("Minutes")] }] }));
    expect(await convertForConversation(Buffer.from(docx), undefined, "m.docx")).toMatchObject({ kind: "text", text: "Minutes" });
    expect(await convertForConversation(Buffer.from("a,b"), "text/csv", "x.csv")).toEqual({ kind: "text", text: "a,b", mimeType: "text/csv" });
  });
});

describe("fileOutput", () => {
  it("puts an image in the result as an image, and only metadata in the JSON", async () => {
    const out = fileOutput({ name: "chart.png" }, await convertForConversation(PNG, "image/png", "chart.png"));
    expect(out.images).toEqual([{ data: PNG.toString("base64"), mimeType: "image/png" }]);
    expect(JSON.stringify(out.data)).not.toContain(PNG.toString("base64"));
  });
});
