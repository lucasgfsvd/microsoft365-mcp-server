import { z } from "zod";

/**
 * Plumbing shared by every Excel surface.
 *
 * All Excel tools target a workbook stored in OneDrive or SharePoint. Callers
 * pass either a driveItem id (preferred) or a drive path, and may pass a
 * workbook session id to chain several edits cheaply.
 */
export const WorkbookRef = z.object({
  driveId: z.string().optional(),
  siteId: z.string().optional(),
  itemId: z.string().describe("driveItem id of the .xlsx file"),
  sessionId: z.string().optional().describe("Workbook session id (pass to chain edits)."),
});

export function drivePrefix(r: { driveId?: string; siteId?: string }): string {
  if (r.driveId) return `/drives/${r.driveId}`;
  if (r.siteId) return `/sites/${r.siteId}/drive`;
  return `/me/drive`;
}

export function withSession(req: ReturnType<import("@microsoft/microsoft-graph-client").Client["api"]>, sessionId?: string) {
  return sessionId ? req.header("workbook-session-id", sessionId) : req;
}

export type Cell = string | number | boolean | null;
