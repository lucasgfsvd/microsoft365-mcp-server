import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { WorkbookRef, drivePrefix, withSession } from "./shared.js";

/** How a sheet looks: charts made from its data, and the formatting of its cells. */

const sheet = (input: { driveId?: string; siteId?: string; itemId: string; worksheet: string }) =>
  `${drivePrefix(input)}/items/${input.itemId}/workbook/worksheets('${encodeURIComponent(input.worksheet)}')`;

const Color = z.string().regex(/^#[0-9a-fA-F]{6}$/, "A colour is #RRGGBB");

export const excelPresentationTools: ToolDefinition[] = [
  {
    name: "excel_add_chart",
    surface: "excel",
    description: "Make a chart from a range of a worksheet (with its header row), placed on that sheet.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All"],
    inputSchema: WorkbookRef.extend({
      worksheet: z.string(),
      sourceAddress: z.string().describe("The data, headers included, e.g. 'A1:C13'."),
      type: z.enum(["ColumnClustered", "ColumnStacked", "BarClustered", "Line", "LineMarkers", "Pie", "Doughnut", "Area", "XYScatter"]).default("ColumnClustered"),
      seriesBy: z.enum(["Auto", "Columns", "Rows"]).default("Auto"),
      title: z.string().optional(),
    }),
    handler: async (input, ctx) => {
      const chart = (await withSession(ctx.graph.api(`${sheet(input)}/charts/add`), input.sessionId).post({
        type: input.type,
        sourceData: input.sourceAddress,
        seriesBy: input.seriesBy,
      })) as { id?: string; name?: string };
      if (input.title && chart.name) {
        await withSession(ctx.graph.api(`${sheet(input)}/charts('${encodeURIComponent(chart.name)}')/title`), input.sessionId).patch({ text: input.title, visible: true });
      }
      return { id: chart.id, name: chart.name, type: input.type, title: input.title };
    },
  },
  {
    name: "excel_format_range",
    surface: "excel",
    description:
      "Format a range: bold, italic, font colour and size, fill colour, number format (e.g. '0.00%', " +
      "'#,##0', 'yyyy-mm-dd'), and fitting column widths to their contents.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All"],
    inputSchema: WorkbookRef.extend({
      worksheet: z.string(),
      address: z.string(),
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      fontColor: Color.optional(),
      fontSize: z.number().min(1).max(409).optional(),
      fillColor: Color.nullable().optional().describe("null clears the fill."),
      numberFormat: z.string().optional().describe("Applied to every cell of the range."),
      autofitColumns: z.boolean().optional(),
    }),
    handler: async (input, ctx) => {
      const range = `${sheet(input)}/range(address='${encodeURIComponent(input.address)}')`;
      const req = (path: string) => withSession(ctx.graph.api(path), input.sessionId);
      const font: Record<string, unknown> = {};
      if (input.bold !== undefined) font.bold = input.bold;
      if (input.italic !== undefined) font.italic = input.italic;
      if (input.fontColor) font.color = input.fontColor;
      if (input.fontSize) font.size = input.fontSize;
      const done: string[] = [];
      if (Object.keys(font).length) {
        await req(`${range}/format/font`).patch(font);
        done.push("font");
      }
      if (input.fillColor === null) {
        await req(`${range}/format/fill/clear`).post({});
        done.push("fill cleared");
      } else if (input.fillColor) {
        await req(`${range}/format/fill`).patch({ color: input.fillColor });
        done.push("fill");
      }
      if (input.numberFormat) {
        // The API takes one format per cell, so the range's shape is needed.
        const shape = (await req(`${range}?$select=rowCount,columnCount`).get()) as { rowCount: number; columnCount: number };
        const grid = Array.from({ length: shape.rowCount }, () => Array.from({ length: shape.columnCount }, () => input.numberFormat));
        await req(range).patch({ numberFormat: grid });
        done.push("number format");
      }
      if (input.autofitColumns) {
        await req(`${range}/format/autofitColumns`).post({});
        done.push("column widths");
      }
      if (!done.length) throw new Error("Give at least one format to apply.");
      return { ok: true, applied: done };
    },
  },
];
