import { describe, it, expect } from "vitest";
import { onenoteTools } from "../src/tools/onenote/index.js";
import { excelTools } from "../src/tools/excel/index.js";
import { allTools } from "../src/tools/index.js";
import { DELEGATED_ONLY } from "../src/tools/appOnly.js";
import { callTool, findTool, makeContext } from "./helpers/mockGraph.js";

describe("onenote_append_to_page", () => {
  it("appends text as escaped paragraphs, one per line", async () => {
    const ctx = makeContext();
    await callTool(findTool(onenoteTools, "onenote_append_to_page"), ctx, { pageId: "p1", text: "Decisions:\nShip <Friday>" });
    expect(ctx.mock.calls[0]).toMatchObject({
      method: "PATCH",
      path: "/me/onenote/pages/p1/content",
      body: [{ target: "body", action: "append", content: "<p>Decisions:</p><p>Ship &lt;Friday&gt;</p>" }],
    });
  });

  it("wants text or html, not both", async () => {
    await expect(callTool(findTool(onenoteTools, "onenote_append_to_page"), makeContext(), { pageId: "p", text: "a", html: "<p>b</p>" })).rejects.toThrow(/text or html/);
  });
});

describe("Excel charts and formatting", () => {
  const sheet = "/me/drive/items/x/workbook/worksheets('Data')";

  it("adds a chart, then titles it", async () => {
    const ctx = makeContext();
    ctx.mock.on("/charts/add", { id: "c1", name: "Chart 1" });
    await callTool(findTool(excelTools, "excel_add_chart"), ctx, { itemId: "x", worksheet: "Data", sourceAddress: "A1:B13", type: "Line", title: "Revenue" });
    expect(ctx.mock.calls.map((c) => [c.method, c.path, c.body])).toEqual([
      ["POST", `${sheet}/charts/add`, { type: "Line", sourceData: "A1:B13", seriesBy: "Auto" }],
      ["PATCH", `${sheet}/charts('Chart%201')/title`, { text: "Revenue", visible: true }],
    ]);
  });

  it("formats font, fill and number format over the whole range", async () => {
    const ctx = makeContext();
    ctx.mock.on("?$select=rowCount,columnCount", { rowCount: 2, columnCount: 3 });
    const out = await callTool(findTool(excelTools, "excel_format_range"), ctx, { itemId: "x", worksheet: "Data", address: "B2:D3", bold: true, fillColor: "#FFF2CC", numberFormat: "0.0%" });
    const range = `${sheet}/range(address='B2%3AD3')`;
    expect(ctx.mock.calls.map((c) => [c.method, c.path])).toEqual([
      ["PATCH", `${range}/format/font`],
      ["PATCH", `${range}/format/fill`],
      ["GET", `${range}?$select=rowCount,columnCount`],
      ["PATCH", range],
    ]);
    expect(ctx.mock.calls[3]!.body).toEqual({ numberFormat: [["0.0%", "0.0%", "0.0%"], ["0.0%", "0.0%", "0.0%"]] });
    expect(out).toEqual({ ok: true, applied: ["font", "fill", "number format"] });
  });

  it("refuses a colour it cannot use, and an empty format", async () => {
    const tool = findTool(excelTools, "excel_format_range");
    await expect(callTool(tool, makeContext(), { itemId: "x", worksheet: "Data", address: "A1", fillColor: "yellow" })).rejects.toThrow(/#RRGGBB/);
    await expect(callTool(tool, makeContext(), { itemId: "x", worksheet: "Data", address: "A1" })).rejects.toThrow(/at least one/);
  });
});

describe("app-only list", () => {
  it("names only tools that exist", () => {
    const names = new Set(allTools().map((t) => t.name));
    expect([...DELEGATED_ONLY].filter((n) => !names.has(n))).toEqual([]);
  });
});
