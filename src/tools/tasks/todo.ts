import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { fetchPage } from "../../graph/pagination.js";
import { PaginationInput, trimEmpty } from "../../util/schema.js";

/** Microsoft To Do: the user's personal task lists. */
export const todoTools: ToolDefinition[] = [
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
  {
    name: "todo_delete_task",
    surface: "tasks",
    description: "Delete a To Do task. It is removed outright; To Do has no recycle bin.",
    mutating: true,
    requiredScopes: ["Tasks.ReadWrite"],
    inputSchema: z.object({ listId: z.string(), taskId: z.string() }),
    handler: async (input, ctx) => {
      await ctx.graph.api(`/me/todo/lists/${input.listId}/tasks/${input.taskId}`).delete();
      return { ok: true };
    },
  },
];
