import { z } from "zod";

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
