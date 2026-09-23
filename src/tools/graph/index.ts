import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { batchGet, MAX_BATCH_SIZE } from "../../graph/batch.js";

export const graphTools: ToolDefinition[] = [
  {
    name: "graph_batch_get",
    surface: "graph",
    // Reads only, so this stays outside the write-guard's remit by construction:
    // there is no way to express a mutation through it.
    description:
      `Run up to ${MAX_BATCH_SIZE} Microsoft Graph GET requests as a single round trip. ` +
      "Use this whenever you need several independent reads — it is dramatically faster " +
      "than calling the individual tools one after another, which is what makes " +
      "multi-surface questions (\"what does my week look like?\") practical.\n\n" +
      "Each request takes an `id` you choose and a Graph-relative `url`, e.g. " +
      '`{"id":"mail","url":"/me/messages?$top=5&$select=subject,from"}`. Useful starting points: ' +
      "`/me/messages`, `/me/events`, `/me/calendarView?startDateTime=…&endDateTime=…`, " +
      "`/me/drive/recent`, `/me/todo/lists`, `/me/joinedTeams`, `/me/people`.\n\n" +
      "Results come back in the order given, each with its own `status`. A sub-request " +
      "that fails does not fail the rest — check `status` per item. GET only; use the " +
      "surface-specific tools to create, update or send anything.",
    inputSchema: z.object({
      requests: z
        .array(
          z.object({
            id: z.string().min(1).describe("Your label for this request; echoed back on the result."),
            url: z
              .string()
              .min(1)
              .describe('Graph-relative URL starting with "/", e.g. "/me/events?$top=10".'),
          }),
        )
        .min(1)
        .max(MAX_BATCH_SIZE)
        .describe(`Between 1 and ${MAX_BATCH_SIZE} GET requests.`),
    }),
    handler: async (input, ctx) => {
      const results = await batchGet(ctx.graph, input.requests);
      return {
        count: results.length,
        succeeded: results.filter((r) => r.status >= 200 && r.status < 300).length,
        responses: results,
      };
    },
  },
];
