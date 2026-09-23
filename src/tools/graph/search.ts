import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { searchQuery, SEARCH_ENTITY_TYPES } from "../../graph/search.js";

export const graphSearchTools: ToolDefinition[] = [
  {
    name: "graph_search",
    surface: "graph",
    description:
      "Search across Microsoft 365 with one relevance-ranked query — mail, files, " +
      "SharePoint, Teams messages or people. Prefer this over the per-surface " +
      "*_search tools when you don't already know where something lives: " +
      '"everything about the Q3 thermal review" is one call here instead of four ' +
      "separate searches whose results you then have to merge.\n\n" +
      "Graph will not mix entity types in one request except within the " +
      "SharePoint/OneDrive family (driveItem, list, listItem, site). message, " +
      "event, chatMessage and person each need their own call — including " +
      "message and event, which cannot be combined. The tool refuses an invalid " +
      "combination before spending a round trip, and names the groups.",
    requiredScopes: ["Mail.Read", "Files.Read.All", "Sites.Read.All"],
    inputSchema: z.object({
      query: z.string().min(1).describe("Free-text query, e.g. 'thermal review Q3'."),
      entityTypes: z
        .array(z.enum(SEARCH_ENTITY_TYPES))
        .min(1)
        .describe("What to search. Combine only driveItem/list/listItem/site; message, event, chatMessage and person must each be alone."),
      size: z.number().int().min(1).max(100).optional().describe("Hits to return (default 25)."),
      from: z.number().int().min(0).optional().describe("Offset for paging (default 0)."),
    }),
    handler: async (input, ctx) =>
      searchQuery(ctx.graph, {
        query: input.query,
        entityTypes: input.entityTypes,
        size: input.size,
        from: input.from,
      }),
  },
];
