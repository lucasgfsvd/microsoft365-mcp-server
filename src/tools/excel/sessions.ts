import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { WorkbookRef, drivePrefix, withSession } from "./shared.js";

/**
 * Workbook session lifecycle, plus forcing a recalculation.
 */

export const excelSessionTools: ToolDefinition[] = [
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
];
