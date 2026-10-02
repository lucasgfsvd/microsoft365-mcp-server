import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { fetchPage } from "../../graph/pagination.js";
import { listPlans } from "../../graph/planner.js";
import { PaginationInput, trimEmpty } from "../../util/schema.js";

/** Planner: the team plans in Microsoft 365 groups and Teams. */
export const plannerTools: ToolDefinition[] = [
  {
    name: "planner_list_plans",
    surface: "tasks",
    description:
      "List the user's Planner plans: those shared with them and those of every team they have " +
      "joined, each marked with its group. Pass groupId to list one Microsoft 365 group's plans " +
      "(plans in groups that are not teams can only be reached that way).",
    requiredScopes: ["Tasks.Read"],
    inputSchema: z.object({
      groupId: z.string().optional().describe("A Microsoft 365 group (or team) id, to list only its plans."),
    }),
    handler: async (input, ctx) => listPlans(ctx.graph, input.groupId),
  },
  {
    name: "planner_list_tasks",
    surface: "tasks",
    description: "List tasks in a Planner plan.",
    requiredScopes: ["Tasks.Read"],
    inputSchema: PaginationInput.extend({ planId: z.string() }),
    handler: async (input, ctx) => fetchPage(ctx.graph, `/planner/plans/${input.planId}/tasks`, input),
  },
  {
    name: "planner_create_task",
    surface: "tasks",
    description: "Create a Planner task.",
    mutating: true,
    requiredScopes: ["Tasks.ReadWrite"],
    inputSchema: z.object({
      planId: z.string(),
      title: z.string(),
      bucketId: z.string().optional(),
      assigneeUserIds: z.array(z.string()).optional(),
      dueDateTime: z.string().optional(),
    }),
    handler: async (input, ctx) => {
      const assignments: Record<string, unknown> = {};
      for (const uid of input.assigneeUserIds ?? []) {
        assignments[uid] = { "@odata.type": "#microsoft.graph.plannerAssignment", orderHint: " !" };
      }
      return ctx.graph.api(`/planner/tasks`).post(
        trimEmpty({
          planId: input.planId,
          title: input.title,
          bucketId: input.bucketId,
          dueDateTime: input.dueDateTime,
          assignments: Object.keys(assignments).length ? assignments : undefined,
        }),
      );
    },
  },
  {
    name: "planner_complete_task",
    surface: "tasks",
    description: "Mark a Planner task as 100% complete.",
    mutating: true,
    requiredScopes: ["Tasks.ReadWrite"],
    inputSchema: z.object({ taskId: z.string(), etag: z.string().describe("The @odata.etag from the current task.") }),
    handler: async (input, ctx) =>
      ctx.graph
        .api(`/planner/tasks/${input.taskId}`)
        .header("If-Match", input.etag)
        .patch({ percentComplete: 100 }),
  },
  {
    name: "planner_delete_task",
    surface: "tasks",
    description:
      "Delete a Planner task, for every member of the plan. Pass the etag you last saw to delete it " +
      "only if nobody has changed it since; without one, the current version is deleted.",
    mutating: true,
    requiredScopes: ["Tasks.ReadWrite"],
    inputSchema: z.object({
      taskId: z.string(),
      etag: z.string().optional().describe("The task's @odata.etag. Omit to delete whatever version is current."),
    }),
    handler: async (input, ctx) => {
      const url = `/planner/tasks/${input.taskId}`;
      // Planner refuses a DELETE without If-Match, so fetch the etag when none was given.
      const etag = input.etag ?? ((await ctx.graph.api(url).select("id").get()) as { "@odata.etag"?: string })["@odata.etag"];
      if (!etag) throw new Error(`Planner returned no etag for task ${input.taskId}.`);
      await ctx.graph.api(url).header("If-Match", etag).delete();
      return { ok: true };
    },
  },
];
