import { randomUUID } from "node:crypto";
import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { etagOf } from "./planner.js";

interface Details {
  description?: string;
  checklist?: Record<string, { title?: string; isChecked?: boolean; orderHint?: string } | null>;
}

const checklistOf = (d: Details) =>
  Object.entries(d.checklist ?? {}).flatMap(([id, c]) => (c ? [{ id, title: c.title, done: c.isChecked === true }] : []));

/** Planner buckets (the board's columns) and a task's notes and checklist. */
export const plannerDetailTools: ToolDefinition[] = [
  {
    name: "planner_list_buckets",
    surface: "tasks",
    description: "List a plan's buckets, the columns of its board, to move tasks between them.",
    requiredScopes: ["Tasks.Read"],
    inputSchema: z.object({ planId: z.string() }),
    handler: async ({ planId }, ctx) => {
      const res = (await ctx.graph.api(`/planner/plans/${planId}/buckets`).get()) as { value?: Array<{ id: string; name?: string }> };
      return (res.value ?? []).map((b) => ({ id: b.id, name: b.name }));
    },
  },
  {
    name: "planner_create_bucket",
    surface: "tasks",
    description: "Add a bucket (a column of the board) to a plan.",
    mutating: true,
    requiredScopes: ["Tasks.ReadWrite"],
    inputSchema: z.object({ planId: z.string(), name: z.string().min(1) }),
    handler: async ({ planId, name }, ctx) => {
      const b = (await ctx.graph.api(`/planner/buckets`).post({ planId, name, orderHint: " !" })) as { id: string; name?: string };
      return { id: b.id, name: b.name };
    },
  },
  {
    name: "planner_get_task_details",
    surface: "tasks",
    description: "A Planner task's notes and checklist.",
    requiredScopes: ["Tasks.Read"],
    inputSchema: z.object({ taskId: z.string() }),
    handler: async ({ taskId }, ctx) => {
      const d = (await ctx.graph.api(`/planner/tasks/${taskId}/details`).get()) as Details;
      return { description: d.description ?? "", checklist: checklistOf(d) };
    },
  },
  {
    name: "planner_update_task_details",
    surface: "tasks",
    description: "Change a Planner task's notes, and its checklist: add items, tick or untick them, remove them.",
    mutating: true,
    requiredScopes: ["Tasks.ReadWrite"],
    inputSchema: z
      .object({
        taskId: z.string(),
        description: z.string().optional().describe("Replaces the notes."),
        addItems: z.array(z.string().min(1)).max(20).optional().describe("Checklist items to add, by title."),
        check: z.array(z.string()).optional().describe("Checklist item ids to tick."),
        uncheck: z.array(z.string()).optional(),
        remove: z.array(z.string()).optional(),
      })
      .refine((t) => [t.description, t.addItems, t.check, t.uncheck, t.remove].some((x) => x !== undefined), {
        message: "Give at least one change",
      }),
    handler: async (input, ctx) => {
      const url = `/planner/tasks/${input.taskId}/details`;
      const item = "microsoft.graph.plannerChecklistItem";
      const checklist: Details["checklist"] = {};
      for (const title of input.addItems ?? []) checklist[randomUUID()] = { "@odata.type": item, title, isChecked: false } as never;
      for (const id of input.check ?? []) checklist[id] = { "@odata.type": item, isChecked: true } as never;
      for (const id of input.uncheck ?? []) checklist[id] = { "@odata.type": item, isChecked: false } as never;
      for (const id of input.remove ?? []) checklist[id] = null;
      const patch: Record<string, unknown> = {};
      if (input.description !== undefined) patch.description = input.description;
      if (Object.keys(checklist).length) patch.checklist = checklist;
      const updated = (await ctx.graph
        .api(url)
        .header("If-Match", await etagOf(ctx.graph, url))
        .header("Prefer", "return=representation")
        .patch(patch)) as Details | undefined;
      return updated ? { description: updated.description ?? "", checklist: checklistOf(updated) } : { ok: true };
    },
  },
];
