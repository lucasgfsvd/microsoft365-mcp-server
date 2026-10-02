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
    name: "todo_update_task",
    surface: "tasks",
    description: "Change a To Do task: title, notes, due date, importance, status, or a reminder.",
    mutating: true,
    requiredScopes: ["Tasks.ReadWrite"],
    inputSchema: z
      .object({
        listId: z.string(),
        taskId: z.string(),
        title: z.string().optional(),
        body: z.string().optional().describe("Replaces the task's notes."),
        dueDate: z.iso.date().nullable().optional().describe("YYYY-MM-DD; null clears it."),
        importance: z.enum(["low", "normal", "high"]).optional(),
        status: z.enum(["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"]).optional(),
        reminder: z.object({ dateTime: z.string(), timeZone: z.string().default("UTC") }).nullable().optional().describe("null removes it."),
      })
      .refine((t) => [t.title, t.body, t.dueDate, t.importance, t.status, t.reminder].some((x) => x !== undefined), {
        message: "Give at least one thing to change",
      }),
    handler: async (input, ctx) => {
      const patch: Record<string, unknown> = {};
      if (input.title !== undefined) patch.title = input.title;
      if (input.body !== undefined) patch.body = { content: input.body, contentType: "text" };
      if (input.dueDate !== undefined) patch.dueDateTime = input.dueDate === null ? null : { dateTime: `${input.dueDate}T00:00:00`, timeZone: "UTC" };
      if (input.importance) patch.importance = input.importance;
      if (input.status) patch.status = input.status;
      if (input.reminder !== undefined) {
        patch.isReminderOn = input.reminder !== null;
        if (input.reminder) patch.reminderDateTime = input.reminder;
      }
      return ctx.graph.api(`/me/todo/lists/${input.listId}/tasks/${input.taskId}`).patch(patch);
    },
  },
  {
    name: "todo_add_checklist_item",
    surface: "tasks",
    description: "Add a step (checklist item) to a To Do task.",
    mutating: true,
    requiredScopes: ["Tasks.ReadWrite"],
    inputSchema: z.object({ listId: z.string(), taskId: z.string(), title: z.string().min(1) }),
    handler: async (input, ctx) =>
      ctx.graph.api(`/me/todo/lists/${input.listId}/tasks/${input.taskId}/checklistItems`).post({ displayName: input.title }),
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
