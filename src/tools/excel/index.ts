import type { ToolDefinition } from "../../types.js";
import { excelSessionTools } from "./sessions.js";
import { excelWorksheetTools } from "./worksheets.js";
import { excelRangeTools } from "./ranges.js";
import { excelTableTools } from "./tables.js";
import { excelWorkbookTools } from "./workbook.js";

/** Every Excel tool, composed from the per-concern modules beside this file. */
export const excelTools: ToolDefinition[] = [
  ...excelSessionTools,
  ...excelWorksheetTools,
  ...excelRangeTools,
  ...excelTableTools,
  ...excelWorkbookTools,
];
