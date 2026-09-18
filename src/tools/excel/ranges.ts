import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { WorkbookRef, drivePrefix, withSession } from "./shared.js";

/**
 * Cell-range reads and writes, including formulas.
 */

export const excelRangeTools: ToolDefinition[] = [
  {
    name: "excel_get_range",
    surface: "excel",
    description: "Read a cell range from a worksheet (values, text, formulas).",
    requiredScopes: ["Files.Read.All"],
    inputSchema: WorkbookRef.extend({
      worksheet: z.string().describe("Worksheet name, e.g. 'Sheet1'"),
      address: z.string().describe("A1 range, e.g. 'A1:D20'"),
    }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      return withSession(
        ctx.graph.api(
          `${base}/items/${input.itemId}/workbook/worksheets('${encodeURIComponent(
            input.worksheet,
          )}')/range(address='${encodeURIComponent(input.address)}')`,
        ),
        input.sessionId,
      ).get();
    },
  },
  {
    name: "excel_update_range",
    surface: "excel",
    description: "Write values to a cell range.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All"],
    inputSchema: WorkbookRef.extend({
      worksheet: z.string(),
      address: z.string(),
      values: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))),
    }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      return withSession(
        ctx.graph.api(
          `${base}/items/${input.itemId}/workbook/worksheets('${encodeURIComponent(
            input.worksheet,
          )}')/range(address='${encodeURIComponent(input.address)}')`,
        ),
        input.sessionId,
      ).patch({ values: input.values });
    },
  },
  {
    name: "excel_set_formula",
    surface: "excel",
    description:
      "Write formulas into a range. Pass a 2D array of formula strings (each starting with '=' or a " +
      "plain value). Shape must match the range.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All"],
    inputSchema: WorkbookRef.extend({
      worksheet: z.string(),
      address: z.string().describe("A1 range, e.g. 'B2' or 'C2:C10'"),
      formulas: z.array(z.array(z.string())),
    }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      return withSession(
        ctx.graph.api(
          `${base}/items/${input.itemId}/workbook/worksheets('${encodeURIComponent(
            input.worksheet,
          )}')/range(address='${encodeURIComponent(input.address)}')`,
        ),
        input.sessionId,
      ).patch({ formulas: input.formulas });
    },
  },
  {
    name: "excel_clear_range",
    surface: "excel",
    description: "Clear a range's contents, formats, or both (default 'All').",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All"],
    inputSchema: WorkbookRef.extend({
      worksheet: z.string(),
      address: z.string(),
      applyTo: z.enum(["All", "Formats", "Contents"]).default("All"),
    }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      await withSession(
        ctx.graph.api(
          `${base}/items/${input.itemId}/workbook/worksheets('${encodeURIComponent(
            input.worksheet,
          )}')/range(address='${encodeURIComponent(input.address)}')/clear`,
        ),
        input.sessionId,
      ).post({ applyTo: input.applyTo });
      return { ok: true };
    },
  },
];
