import { z } from "zod";
import ExcelJS from "exceljs";
import type { ToolDefinition, ToolContext } from "../../types.js";
import { DrivePath, Filename } from "../../util/schema.js";
import { drivePrefix } from "./shared.js";
import type { Cell } from "./shared.js";
import { uploadContent } from "../../graph/upload.js";

/**
 * Whole-file operations. The only Excel tools that build bytes locally with
 * ExcelJS and upload them, rather than driving the Graph Excel REST API
 * against a workbook that already exists.
 */

async function uploadNewXlsxToPath(
  ctx: ToolContext,
  scope: { driveId?: string; siteId?: string; parentPath: string; filename: string },
  buf: Buffer,
): Promise<unknown> {
  return uploadContent(ctx.graph, drivePrefix(scope), scope, buf);
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

export const excelWorkbookTools: ToolDefinition[] = [
  {
    name: "excel_create_workbook",
    surface: "excel",
    description:
      "Create a brand new .xlsx with an initial worksheet and optional seeded values, then upload it. " +
      "Returns the driveItem metadata of the new file so follow-up tools can target it.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: z.object({
      driveId: z.string().optional(),
      siteId: z.string().optional(),
      parentPath: DrivePath.describe("Parent folder path, e.g. '/Spreadsheets'"),
      filename: Filename.describe("Filename including .xlsx extension"),
      worksheetName: z.string().default("Sheet1"),
      values: z
        .array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])))
        .optional()
        .describe("Optional 2D array of initial values starting at A1."),
    }),
    handler: async (input, ctx) => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet(input.worksheetName);
      if (input.values && input.values.length) {
        for (let r = 0; r < input.values.length; r++) {
          const row = input.values[r]!;
          for (let c = 0; c < row.length; c++) {
            const cell = row[c] as Cell;
            if (cell !== null) ws.getCell(r + 1, c + 1).value = cell;
          }
        }
      }
      const ab = await wb.xlsx.writeBuffer();
      const buf = Buffer.from(ab as ArrayBuffer);
      const filename = input.filename.endsWith(".xlsx") ? input.filename : `${input.filename}.xlsx`;
      const result = await uploadNewXlsxToPath(
        ctx,
        { driveId: input.driveId, siteId: input.siteId, parentPath: input.parentPath, filename },
        buf,
      );
      return { ok: true, driveItem: result };
    },
  },
  {
    name: "excel_create_from_template",
    surface: "excel",
    description:
      "Create a new .xlsx by copying a template workbook to a new location (and optional new filename). " +
      "Preserves all sheets, formulas, named ranges, tables, and formatting from the template. " +
      "Follow up with excel_update_range / excel_set_formula / excel_add_table_rows to fill in values.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: z
      .object({
        templateDriveId: z.string().optional(),
        templateSiteId: z.string().optional(),
        templateItemId: z.string().optional(),
        templatePath: DrivePath.optional().describe("Template path, e.g. '/Templates/budget.xlsx'"),
        driveId: z.string().optional(),
        siteId: z.string().optional(),
        parentPath: DrivePath.describe("Destination folder path"),
        filename: Filename.describe("New filename, including .xlsx"),
      })
      .refine((d) => Boolean(d.templateItemId) || Boolean(d.templatePath), {
        message: "Provide templateItemId or templatePath",
      }),
    handler: async (input, ctx) => {
      const buf = await downloadTemplateBytes(ctx, input);
      const filename = input.filename.endsWith(".xlsx") ? input.filename : `${input.filename}.xlsx`;
      const result = await uploadNewXlsxToPath(
        ctx,
        { driveId: input.driveId, siteId: input.siteId, parentPath: input.parentPath, filename },
        buf,
      );
      return { ok: true, driveItem: result };
    },
  },
];
