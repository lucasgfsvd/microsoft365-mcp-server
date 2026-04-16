import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { fetchPage } from "../../graph/pagination.js";
import { PaginationInput, trimEmpty } from "../../util/schema.js";

export const tasksTools: ToolDefinition[] = [
  // ---- Microsoft To Do ----
  {
    name: "todo_list_lists",
    surface: "tasks",
    description: "List Microsoft To Do task lists.",
    requiredScopes: ["Tasks.Read"],
    inputSchema: PaginationInput,
    handler: async (input, ctx) => fetchPage(ctx.graph, `/me/todo/lists`, input),
  },
  {
    name: "todo_list_tasks",
    surface: "tasks",
    description: "List tasks in a To Do list.",
    requiredScopes: ["Tasks.Read"],
    inputSchema: PaginationInput.extend({ listId: z.string() }),
    handler: async (input, ctx) => fetchPage(ctx.graph, `/me/todo/lists/${input.listId}/tasks`, input),
  },
  {
    name: "todo_create_task",
    surface: "tasks",
    description: "Create a To Do task.",
    mutating: true,
    requiredScopes: ["Tasks.ReadWrite"],
    inputSchema: z.object({
      listId: z.string(),
      title: z.string(),
      body: z.string().optional(),
      dueDateTime: z.string().optional(),
      importance: z.enum(["low", "normal", "high"]).optional(),
    }),
    handler: async (input, ctx) =>
      ctx.graph.api(`/me/todo/lists/${input.listId}/tasks`).post(
        trimEmpty({
          title: input.title,
          body: input.body ? { content: input.body, contentType: "text" } : undefined,
          dueDateTime: input.dueDateTime ? { dateTime: input.dueDateTime, timeZone: "UTC" } : undefined,
          importance: input.importance,
        }),
      ),
  },
  {
    name: "todo_complete_task",
    surface: "tasks",
    description: "Mark a To Do task as completed.",
    mutating: true,
    requiredScopes: ["Tasks.ReadWrite"],
    inputSchema: z.object({ listId: z.string(), taskId: z.string() }),
    handler: async (input, ctx) =>
      ctx.graph
        .api(`/me/todo/lists/${input.listId}/tasks/${input.taskId}`)
        .patch({ status: "completed" }),
  },

  // ---- Planner ----
  {
    name: "planner_list_plans",
    surface: "tasks",
    description: "List the user's Planner plans.",
    requiredScopes: ["Tasks.Read"],
    inputSchema: PaginationInput,
    handler: async (input, ctx) => fetchPage(ctx.graph, `/me/planner/plans`, input),
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
];
