import { describe, it, expect } from "vitest";
import ExcelJS from "exceljs";
import { cellText, csvField, workbookText } from "../src/ooxml/xlsxText.js";
import { fileContent } from "../src/resources/drive.js";

async function workbook(build: (wb: ExcelJS.Workbook) => void): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  build(wb);
  return Buffer.from(await wb.xlsx.writeBuffer());
}

describe("workbookText", () => {
  it("writes each sheet as CSV under its name, cells where Excel has them", async () => {
    const bytes = await workbook((wb) => {
      const ws = wb.addWorksheet("Pipeline");
      ws.getCell("A1").value = "Deal";
      ws.getCell("B1").value = "Amount";
      ws.getCell("A2").value = "Acme, Inc.";
      ws.getCell("B2").value = 1200;
      ws.getCell("A4").value = 'Say "hi"';
      ws.getCell("C4").value = { formula: "B2*2", result: 2400 };
      wb.addWorksheet("Notes").getCell("B2").value = "second sheet";
    });
    expect(await workbookText(bytes)).toBe(
      [
        "--- Sheet: Pipeline ---",
        "Deal,Amount,",
        '"Acme, Inc.",1200,',
        ",,",
        '"Say ""hi""",,2400',
        "",
        "--- Sheet: Notes ---",
        ",",
        ",second sheet",
      ].join("\n"),
    );
  });

  it("marks hidden sheets, and leaves empty ones empty", async () => {
    const bytes = await workbook((wb) => {
      wb.addWorksheet("Visible").getCell("A1").value = "x";
      wb.addWorksheet("Lookup", { state: "hidden" }).getCell("A1").value = "y";
      wb.addWorksheet("Blank");
    });
    expect(await workbookText(bytes)).toBe("--- Sheet: Visible ---\nx\n\n--- Sheet: Lookup (hidden) ---\ny\n\n--- Sheet: Blank ---\n");
  });

  it("stops at the limit on a whole row, and says so", async () => {
    const bytes = await workbook((wb) => {
      const ws = wb.addWorksheet("Big");
      for (let r = 1; r <= 100; r++) ws.getCell(`A${r}`).value = `row ${r}`;
      wb.addWorksheet("After").getCell("A1").value = "never reached";
    });
    const text = await workbookText(bytes, 60);
    const lines = text.split("\n");
    expect(lines[0]).toBe("--- Sheet: Big ---");
    expect(lines.slice(1, -2).every((l) => /^row \d+$/.test(l))).toBe(true);
    expect(lines.at(-1)).toMatch(/^--- Truncated at 60 characters/);
    expect(text).not.toContain("never reached");
  });
});

describe("cellText", () => {
  it("reads every kind of cell value as it displays", () => {
    expect(cellText(null)).toBe("");
    expect(cellText(true)).toBe("true");
    expect(cellText(new Date("2026-09-28T00:00:00Z"))).toBe("2026-09-28");
    expect(cellText(new Date("2026-09-28T08:30:00Z"))).toBe("2026-09-28T08:30:00.000Z");
    expect(cellText({ richText: [{ text: "bold" }, { text: " plain" }] })).toBe("bold plain");
    expect(cellText({ text: "site", hyperlink: "https://example.com" })).toBe("site");
    expect(cellText({ error: "#DIV/0!" })).toBe("#DIV/0!");
    expect(cellText({ formula: "A1", result: undefined } as unknown as ExcelJS.CellValue)).toBe("");
  });

  it("quotes only the fields that need it", () => {
    expect(csvField("plain")).toBe("plain");
    expect(csvField("a\nb")).toBe('"a\nb"');
  });
});

describe("drive resource", () => {
  it("returns an attached spreadsheet as text, by type or by name", async () => {
    const bytes = await workbook((wb) => (wb.addWorksheet("S").getCell("A1").value = "hello"));
    const xlsx = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    expect(await fileContent("u", bytes, xlsx, "book")).toEqual({ uri: "u", mimeType: "text/plain", text: "--- Sheet: S ---\nhello" });
    expect(await fileContent("u", bytes, undefined, "Book.XLSX")).toMatchObject({ text: "--- Sheet: S ---\nhello" });
  });
});
