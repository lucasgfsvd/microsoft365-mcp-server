import { describe, it, expect } from "vitest";
import { tasksTools } from "../src/tools/tasks/index.js";
import { callTool, findTool, makeContext } from "./helpers/mockGraph.js";

describe("todo_delete_task", () => {
  it("DELETEs the task inside its list", async () => {
    const ctx = makeContext();
    const out = await callTool(findTool(tasksTools, "todo_delete_task"), ctx, { listId: "L1", taskId: "T1" });
    expect(ctx.mock.calls).toEqual([expect.objectContaining({ method: "DELETE", path: "/me/todo/lists/L1/tasks/T1" })]);
    expect(out).toEqual({ ok: true });
  });
});

describe("planner_delete_task", () => {
  const tool = findTool(tasksTools, "planner_delete_task");

  it("deletes only the version the caller saw, when given its etag", async () => {
    const ctx = makeContext();
    await callTool(tool, ctx, { taskId: "T1", etag: 'W/"seen"' });
    expect(ctx.mock.calls).toHaveLength(1);
    expect(ctx.mock.calls[0]).toMatchObject({ method: "DELETE", path: "/planner/tasks/T1", headers: { "If-Match": 'W/"seen"' } });
  });

  it("reads the current etag first when none is given, since Planner requires If-Match", async () => {
    const ctx = makeContext();
    ctx.mock.on("/planner/tasks/T1", { id: "T1", "@odata.etag": 'W/"current"' });
    await callTool(tool, ctx, { taskId: "T1" });
    expect(ctx.mock.calls.map((c) => c.method)).toEqual(["GET", "DELETE"]);
    expect(ctx.mock.calls[1]!.headers["If-Match"]).toBe('W/"current"');
  });

  it("refuses to send a DELETE with no etag at all", async () => {
    const ctx = makeContext();
    ctx.mock.on("/planner/tasks/T1", { id: "T1" });
    await expect(callTool(tool, ctx, { taskId: "T1" })).rejects.toThrow(/no etag/);
    expect(ctx.mock.calls.some((c) => c.method === "DELETE")).toBe(false);
  });
});

describe("planner_update_task", () => {
  const tool = findTool(tasksTools, "planner_update_task");

  it("assigns by email, unassigns, and maps priority, under the task's etag", async () => {
    const ctx = makeContext();
    ctx.mock.on("/$batch", { responses: [{ id: "0", status: 200, body: { id: "u-ana" } }] });
    ctx.mock.on("/planner/tasks/T1", { id: "T1", "@odata.etag": 'W/"now"' });
    await callTool(tool, ctx, { taskId: "T1", assign: ["ana@x.com"], unassign: ["11111111-2222-3333-4444-555555555555"], priority: "urgent", percentComplete: 50 });
    const patch = ctx.mock.calls.find((c) => c.method === "PATCH")!;
    expect(patch.headers["If-Match"]).toBe('W/"now"');
    expect(patch.body).toEqual({
      percentComplete: 50,
      priority: 1,
      assignments: {
        "u-ana": { "@odata.type": "#microsoft.graph.plannerAssignment", orderHint: " !" },
        "11111111-2222-3333-4444-555555555555": null,
      },
    });
  });

  it("refuses an empty change, and an address nobody has", async () => {
    const ctx = makeContext();
    await expect(callTool(tool, ctx, { taskId: "T1" })).rejects.toThrow(/at least one/);
    ctx.mock.on("/$batch", { responses: [{ id: "0", status: 404, body: { error: { message: "Resource not found" } } }] });
    await expect(callTool(tool, ctx, { taskId: "T1", assign: ["ghost@x.com"] })).rejects.toThrow(/ghost@x.com/);
    expect(ctx.mock.calls.some((c) => c.method === "PATCH")).toBe(false);
  });
});

describe("planner task details", () => {
  it("adds, ticks and removes checklist items in one change", async () => {
    const ctx = makeContext();
    ctx.mock.on("/details", { "@odata.etag": 'W/"d1"' });
    ctx.mock.on("/details", { description: "Notes", checklist: { a: { title: "Draft", isChecked: true }, gone: null } });
    const out = await callTool(findTool(tasksTools, "planner_update_task_details"), ctx, { taskId: "T1", addItems: ["Review"], check: ["a"], remove: ["old"] });
    const patch = ctx.mock.calls[1]!;
    expect(patch.headers["If-Match"]).toBe('W/"d1"');
    const checklist = (patch.body as { checklist: Record<string, unknown> }).checklist;
    expect(Object.values(checklist)).toContainEqual({ "@odata.type": "microsoft.graph.plannerChecklistItem", title: "Review", isChecked: false });
    expect(checklist).toMatchObject({ a: { isChecked: true }, old: null });
    expect(out).toEqual({ description: "Notes", checklist: [{ id: "a", title: "Draft", done: true }] });
  });
});

describe("todo_update_task", () => {
  it("sets a due date, clears a reminder, changes status", async () => {
    const ctx = makeContext();
    await callTool(findTool(tasksTools, "todo_update_task"), ctx, { listId: "L", taskId: "T", dueDate: "2026-10-09", reminder: null, status: "inProgress" });
    expect(ctx.mock.calls[0]).toMatchObject({
      method: "PATCH",
      path: "/me/todo/lists/L/tasks/T",
      body: { dueDateTime: { dateTime: "2026-10-09T00:00:00", timeZone: "UTC" }, isReminderOn: false, status: "inProgress" },
    });
  });

  it("adds a step", async () => {
    const ctx = makeContext();
    await callTool(findTool(tasksTools, "todo_add_checklist_item"), ctx, { listId: "L", taskId: "T", title: "Call Bob" });
    expect(ctx.mock.calls[0]).toMatchObject({ path: "/me/todo/lists/L/tasks/T/checklistItems", body: { displayName: "Call Bob" } });
  });
});
