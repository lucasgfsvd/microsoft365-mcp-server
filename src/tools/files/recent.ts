import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { recentFiles } from "../../graph/recentFiles.js";

export const recentFileTools: ToolDefinition[] = [
  {
    name: "files_list_recent",
    surface: "files",
    description:
      "The documents you can reach that changed most recently, across OneDrive and SharePoint, newest first, " +
      "with who changed them. (Microsoft is retiring its 'recent' and 'shared with me' lists; this uses search.)",
    requiredScopes: ["Files.Read.All", "Sites.Read.All"],
    inputSchema: z.object({ top: z.number().int().min(1).max(50).default(20) }),
    handler: async ({ top }, ctx) => recentFiles(ctx.graph, top),
  },
];
