import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { WorkbookRef, drivePrefix, withSession } from "./shared.js";

/**
 * Worksheet-level operations: list, add, delete, rename.
 */

export const excelWorksheetTools: ToolDefinition[] = [
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
];
