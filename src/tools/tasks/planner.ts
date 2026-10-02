import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { fetchPage } from "../../graph/pagination.js";
import { listPlans } from "../../graph/planner.js";
import { PaginationInput, trimEmpty } from "../../util/schema.js";
import { resolveUserIds } from "../../graph/users.js";
import type { ToolContext } from "../../types.js";

/**
 * Planner refuses a change without If-Match. The caller's etag protects their
 * view of the item; without one, the current version's is fetched.
 */
export async function etagOf(graph: ToolContext["graph"], url: string, given?: string): Promise<string> {
  if (given) return given;
  const etag = ((await graph.api(url).get()) as { "@odata.etag"?: string })["@odata.etag"];
  if (!etag) throw new Error(`Planner returned no etag for ${url}.`);
  return etag;
}

const assignment = { "@odata.type": "#microsoft.graph.plannerAssignment", orderHint: " !" };
const People = z.array(z.string().min(3)).max(20).describe("Directory ids or email addresses.");

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
      assigneeUserIds: People.optional(),
      dueDateTime: z.string().optional(),
    }),
    handler: async (input, ctx) => {
      const assignments: Record<string, unknown> = {};
      for (const uid of await resolveUserIds(ctx.graph, input.assigneeUserIds ?? [])) assignments[uid] = assignment;
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
      await ctx.graph.api(url).header("If-Match", await etagOf(ctx.graph, url, input.etag)).delete();
      return { ok: true };
    },
  },
  {
    name: "planner_update_task",
    surface: "tasks",
    description:
      "Change a Planner task: title, due date, progress (0, 50 or 100), priority, its bucket, and who is " +
      "assigned (by email or id). Pass the etag you last saw to change it only if nobody else has since.",
    mutating: true,
    requiredScopes: ["Tasks.ReadWrite", "User.ReadBasic.All"],
    inputSchema: z
      .object({
        taskId: z.string(),
        etag: z.string().optional(),
        title: z.string().optional(),
        dueDateTime: z.string().nullable().optional().describe("ISO date-time; null clears it."),
        percentComplete: z.union([z.literal(0), z.literal(50), z.literal(100)]).optional(),
        priority: z.enum(["urgent", "important", "medium", "low"]).optional(),
        bucketId: z.string().optional(),
        assign: People.optional(),
        unassign: People.optional(),
      })
      .refine((t) => Object.keys(t).some((k) => !["taskId", "etag"].includes(k) && t[k as keyof typeof t] !== undefined), {
        message: "Give at least one thing to change",
      }),
    handler: async (input, ctx) => {
      const url = `/planner/tasks/${input.taskId}`;
      const patch: Record<string, unknown> = {};
      if (input.title !== undefined) patch.title = input.title;
      if (input.dueDateTime !== undefined) patch.dueDateTime = input.dueDateTime;
      if (input.percentComplete !== undefined) patch.percentComplete = input.percentComplete;
      if (input.priority) patch.priority = { urgent: 1, important: 3, medium: 5, low: 9 }[input.priority as string];
      if (input.bucketId) patch.bucketId = input.bucketId;
      const assignments: Record<string, unknown> = {};
      for (const uid of await resolveUserIds(ctx.graph, input.assign ?? [])) assignments[uid] = assignment;
      for (const uid of await resolveUserIds(ctx.graph, input.unassign ?? [])) assignments[uid] = null;
      if (Object.keys(assignments).length) patch.assignments = assignments;
      const etag = await etagOf(ctx.graph, url, input.etag);
      // Planner answers 204 unless asked for the updated task.
      return (await ctx.graph.api(url).header("If-Match", etag).header("Prefer", "return=representation").patch(patch)) ?? { ok: true };
    },
  },
];
