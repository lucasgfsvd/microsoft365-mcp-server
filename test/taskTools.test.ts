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
