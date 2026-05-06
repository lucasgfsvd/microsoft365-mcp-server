import { z } from "zod";
import ExcelJS from "exceljs";
import type { ToolDefinition, ToolContext } from "../../types.js";
import { DrivePath, Filename } from "../../util/schema.js";

/**
 * All Excel tools target a workbook stored in OneDrive or SharePoint.
 * Callers pass either a driveItem id (preferred) or a drive path. We use persistent
 * workbook sessions for performance when making multiple edits.
 *
 * excel_create_workbook generates a new .xlsx via ExcelJS and uploads it; every other
 * tool operates on an existing workbook via the Microsoft Graph Excel REST API.
 */
const WorkbookRef = z.object({
  driveId: z.string().optional(),
  siteId: z.string().optional(),
  itemId: z.string().describe("driveItem id of the .xlsx file"),
  sessionId: z.string().optional().describe("Workbook session id (pass to chain edits)."),
});

function drivePrefix(r: { driveId?: string; siteId?: string }): string {
  if (r.driveId) return `/drives/${r.driveId}`;
  if (r.siteId) return `/sites/${r.siteId}/drive`;
  return `/me/drive`;
}

function withSession(req: ReturnType<import("@microsoft/microsoft-graph-client").Client["api"]>, sessionId?: string) {
  return sessionId ? req.header("workbook-session-id", sessionId) : req;
}

async function uploadNewXlsxToPath(
  ctx: ToolContext,
  scope: { driveId?: string; siteId?: string; parentPath: string; filename: string },
  buf: Buffer,
): Promise<unknown> {
  const base = drivePrefix(scope);
  const parent = scope.parentPath.replace(/\/$/, "");
  const path = `${base}/root:${parent}/${scope.filename}:/content`;
  return ctx.graph.api(path).put(buf);
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

type Cell = string | number | boolean | null;

export const excelTools: ToolDefinition[] = [
  {
    name: "excel_create_session",
    surface: "excel",
    description: "Create a persistent workbook session (improves throughput for multi-step edits).",
    requiredScopes: ["Files.ReadWrite.All"],
    inputSchema: WorkbookRef.extend({
      persistChanges: z.boolean().default(true),
    }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      return ctx.graph
        .api(`${base}/items/${input.itemId}/workbook/createSession`)
        .post({ persistChanges: input.persistChanges });
    },
  },
  {
    name: "excel_close_session",
    surface: "excel",
    description: "Close a workbook session.",
    requiredScopes: ["Files.ReadWrite.All"],
    inputSchema: WorkbookRef.extend({ sessionId: z.string() }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      await withSession(
        ctx.graph.api(`${base}/items/${input.itemId}/workbook/closeSession`),
        input.sessionId,
      ).post({});
      return { ok: true };
    },
  },
  {
    name: "excel_list_worksheets",
    surface: "excel",
    description: "List worksheets in a workbook.",
    requiredScopes: ["Files.Read.All"],
    inputSchema: WorkbookRef,
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      return withSession(ctx.graph.api(`${base}/items/${input.itemId}/workbook/worksheets`), input.sessionId).get();
    },
  },
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
    name: "excel_list_tables",
    surface: "excel",
    description: "List workbook tables.",
    requiredScopes: ["Files.Read.All"],
    inputSchema: WorkbookRef,
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      return withSession(
        ctx.graph.api(`${base}/items/${input.itemId}/workbook/tables`),
        input.sessionId,
      ).get();
    },
  },
  {
    name: "excel_get_table_rows",
    surface: "excel",
    description: "Read rows of a named table.",
    requiredScopes: ["Files.Read.All"],
    inputSchema: WorkbookRef.extend({ tableName: z.string() }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      return withSession(
        ctx.graph.api(
          `${base}/items/${input.itemId}/workbook/tables('${encodeURIComponent(input.tableName)}')/rows`,
        ),
        input.sessionId,
      ).get();
    },
  },
  {
    name: "excel_add_table_rows",
    surface: "excel",
    description: "Append rows to a named table.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All"],
    inputSchema: WorkbookRef.extend({
      tableName: z.string(),
      values: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))),
    }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      return withSession(
        ctx.graph.api(
          `${base}/items/${input.itemId}/workbook/tables('${encodeURIComponent(input.tableName)}')/rows/add`,
        ),
        input.sessionId,
      ).post({ values: input.values });
    },
  },
  {
    name: "excel_run_workbook_calculation",
    surface: "excel",
    description: "Force the workbook to recalculate.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All"],
    inputSchema: WorkbookRef.extend({ calculationType: z.enum(["Recalculate", "Full", "FullRebuild"]).default("Recalculate") }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      await withSession(
        ctx.graph.api(`${base}/items/${input.itemId}/workbook/application/calculate`),
        input.sessionId,
      ).post({ calculationType: input.calculationType });
      return { ok: true };
    },
  },
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
  {
    name: "excel_add_worksheet",
    surface: "excel",
    description: "Add a new worksheet to an existing workbook.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All"],
    inputSchema: WorkbookRef.extend({ name: z.string() }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      return withSession(
        ctx.graph.api(`${base}/items/${input.itemId}/workbook/worksheets/add`),
        input.sessionId,
      ).post({ name: input.name });
    },
  },
  {
    name: "excel_delete_worksheet",
    surface: "excel",
    description: "Delete a worksheet by name.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All"],
    inputSchema: WorkbookRef.extend({ worksheet: z.string() }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      await withSession(
        ctx.graph.api(
          `${base}/items/${input.itemId}/workbook/worksheets('${encodeURIComponent(input.worksheet)}')`,
        ),
        input.sessionId,
      ).delete();
      return { ok: true };
    },
  },
  {
    name: "excel_rename_worksheet",
    surface: "excel",
    description: "Rename a worksheet.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All"],
    inputSchema: WorkbookRef.extend({ worksheet: z.string(), newName: z.string() }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      return withSession(
        ctx.graph.api(
          `${base}/items/${input.itemId}/workbook/worksheets('${encodeURIComponent(input.worksheet)}')`,
        ),
        input.sessionId,
      ).patch({ name: input.newName });
    },
  },
  {
    name: "excel_create_table",
    surface: "excel",
    description:
      "Create a table from a cell range. With hasHeaders=true the first row is used as column names. " +
      "Use the returned table name with excel_add_table_rows / excel_get_table_rows.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All"],
    inputSchema: WorkbookRef.extend({
      address: z.string().describe("Full address including sheet, e.g. 'Sheet1!A1:D20'"),
      hasHeaders: z.boolean().default(true),
      tableName: z.string().optional().describe("Optional rename of the table after creation."),
    }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      const created = await withSession(
        ctx.graph.api(`${base}/items/${input.itemId}/workbook/tables/add`),
        input.sessionId,
      ).post({ address: input.address, hasHeaders: input.hasHeaders });
      if (input.tableName) {
        const createdName = (created as { name?: string }).name;
        if (createdName) {
          await withSession(
            ctx.graph.api(
              `${base}/items/${input.itemId}/workbook/tables('${encodeURIComponent(createdName)}')`,
            ),
            input.sessionId,
          ).patch({ name: input.tableName });
        }
      }
      return { ok: true, table: created, renamedTo: input.tableName };
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
