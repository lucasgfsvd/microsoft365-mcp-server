import ExcelJS from "exceljs";

/**
 * Most text a workbook may turn into. A small .xlsx can describe a very large
 * sheet, and everything here lands in the conversation.
 */
export const XLSX_TEXT_LIMIT = 1_000_000;

type CellValue = ExcelJS.CellValue;

/** A cell as it reads: formulas by their result, rich text and links by their text. */
export function cellText(value: CellValue): string {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) {
    const iso = value.toISOString();
    return iso.endsWith("T00:00:00.000Z") ? iso.slice(0, 10) : iso;
  }
  if (typeof value !== "object") return String(value);
  if ("richText" in value) return value.richText.map((r) => r.text).join("");
  if ("formula" in value || "sharedFormula" in value) return cellText((value as { result?: CellValue }).result ?? null);
  if ("error" in value) return String(value.error);
  if ("text" in value) return cellText((value as { text: CellValue }).text);
  return "";
}

/** One CSV field: quoted when it holds a comma, quote or line break. */
export function csvField(s: string): string {
  return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

/**
 * Rows of a sheet as CSV, from A1 to its last non-empty cell, so a line and
 * column here are the row and column of the same cell in Excel.
 */
function sheetCsv(ws: ExcelJS.Worksheet): string {
  const rows: string[][] = [];
  let width = 0;
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    const cells: string[] = [];
    row.eachCell({ includeEmpty: false }, (cell, col) => {
      cells[col - 1] = cellText(cell.value);
    });
    let last = cells.length;
    while (last > 0 && !cells[last - 1]) last--;
    if (last === 0) return;
    rows[rowNumber - 1] = cells.slice(0, last);
    width = Math.max(width, last);
  });
  const lines: string[] = [];
  for (let i = 0; i < rows.length; i++) {
    const cells = rows[i] ?? [];
    lines.push(Array.from({ length: width }, (_, c) => csvField(cells[c] ?? "")).join(","));
  }
  return lines.join("\n");
}

/**
 * A workbook as text: every sheet as CSV under a heading naming it. Formulas
 * show their last calculated value, as saved in the file. Stops at `limit`
 * characters and says so, rather than returning a fraction silently.
 */
export async function workbookText(bytes: Buffer, limit = XLSX_TEXT_LIMIT): Promise<string> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes as unknown as Parameters<typeof wb.xlsx.load>[0]);
  const parts: string[] = [];
  let size = 0;
  for (const ws of wb.worksheets) {
    const hidden = ws.state && ws.state !== "visible" ? " (hidden)" : "";
    const part = `--- Sheet: ${ws.name}${hidden} ---\n${sheetCsv(ws)}`;
    if (size + part.length > limit) {
      // Cut at a line break, so no row is left half-written.
      const cut = part.lastIndexOf("\n", limit - size);
      if (cut > 0 && part.indexOf("\n") < cut) parts.push(part.slice(0, cut));
      parts.push(`--- Truncated at ${limit} characters: read the rest with excel_get_range ---`);
      break;
    }
    parts.push(part);
    size += part.length + 2;
  }
  return parts.join("\n\n");
}
