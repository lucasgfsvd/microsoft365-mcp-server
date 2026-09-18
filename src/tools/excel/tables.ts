import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { WorkbookRef, drivePrefix, withSession } from "./shared.js";

/**
 * Excel table (ListObject) operations.
 */

export const excelTableTools: ToolDefinition[] = [
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
];
