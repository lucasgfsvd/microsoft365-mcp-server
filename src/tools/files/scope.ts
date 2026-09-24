import { z } from "zod";

/**
 * Graph address of a drive item given its path. The root has to be `root`:
 * the path form `root:/:` that "/" would otherwise produce is rejected by Graph
 * ("Resource not found for the segment 'root:'", seen live). Append "/children",
 * "/content" and so on as needed.
 */
export function itemByPath(base: string, path: string | undefined): string {
  const p = (path ?? "").replace(/\/+$/, "");
  return p === "" ? `${base}/root` : `${base}/root:${p.startsWith("/") ? p : `/${p}`}:`;
}

/** Graph prefix for the drive a files tool acts on. */
export function drivePrefix(input: { driveId?: string; siteId?: string }): string {
  if (input.driveId) return `/drives/${input.driveId}`;
  if (input.siteId) return `/sites/${input.siteId}/drive`;
  return `/me/drive`;
}

export const ScopeInput = z.object({
  driveId: z.string().optional().describe("Drive id. Omit to use the user's OneDrive."),
  siteId: z.string().optional().describe("SharePoint site id (use sites_search to discover)."),
});
