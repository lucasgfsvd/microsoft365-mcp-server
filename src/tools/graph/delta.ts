import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { DELTA_RESOURCES, fetchDelta } from "../../graph/delta.js";

export const graphDeltaTools: ToolDefinition[] = [
  {
    name: "graph_delta",
    surface: "graph",
    description:
      "Incremental sync: what changed since last time — additions, modifications and " +
      "deletions. Use it for anything recurring (\"what's new in my inbox since this " +
      "morning?\", \"which files changed?\"), and whenever you need to know what was " +
      "*deleted*: a normal listing can only show that something is absent.\n\n" +
      "First call: pass `resource` (mail | calendar | drive | contacts | todo) plus what " +
      "it needs — mail takes an optional `folderId` (default inbox), calendar needs " +
      "`startDateTime` and `endDateTime`, todo needs `listId`. That call returns the " +
      "current state and a `deltaLink`.\n\n" +
      "Later calls: pass that `deltaLink` back as `cursor` and nothing else; you get only " +
      "what changed since. Keep the link yourself — the server stores nothing.\n\n" +
      "Deletions arrive in `removed` as {id, reason}. If `complete` is false the result " +
      "was cut off at `maxItems`: pass `nextLink` back as `cursor` to continue, and only " +
      "treat the sync as done once a `deltaLink` comes back. Drive may list an item more " +
      "than once in one sync; the last occurrence wins.",
    requiredScopes: ["Mail.Read", "Calendars.Read", "Files.Read", "Contacts.Read", "Tasks.Read"],
    inputSchema: z.object({
      cursor: z
        .string()
        .optional()
        .describe("deltaLink or nextLink from a previous graph_delta call. When given, omit everything else."),
      resource: z.enum(DELTA_RESOURCES).optional().describe("What to sync, on the first call."),
      folderId: z.string().optional().describe("mail: folder id or well-known name (default 'inbox')."),
      listId: z.string().optional().describe("todo: the task list id."),
      startDateTime: z.string().optional().describe("calendar: window start, ISO 8601."),
      endDateTime: z.string().optional().describe("calendar: window end, ISO 8601."),
      select: z
        .array(z.string())
        .optional()
        .describe("Fields to return (mail, drive, contacts only). Mail defaults to a summary without bodies."),
      maxItems: z
        .number()
        .int()
        .min(1)
        .max(5000)
        .optional()
        .describe("Stop after about this many items and hand back a nextLink (default 500). Drive pages may overshoot slightly."),
    }),
    handler: async (input, ctx) => {
      const { cursor, resource, maxItems, ...params } = input;
      if (cursor) return fetchDelta(ctx.graph, { cursor }, { maxItems });
      if (!resource) throw new Error("Pass either `cursor` from a previous call, or `resource` to start a sync.");
      return fetchDelta(ctx.graph, { resource, params }, { maxItems });
    },
  },
];
