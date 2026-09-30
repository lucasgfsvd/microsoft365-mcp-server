import { describe, it, expect } from "vitest";
import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { listPlans } from "../src/graph/planner.js";

/** A $batch endpoint answering each sub-request from a url → [status, body] table. */
function batchGraph(table: Record<string, [number, unknown]>) {
  const batches: string[][] = [];
  const graph = {
    api: () => ({
      post: async (payload: { requests: Array<{ id: string; url: string }> }) => {
        batches.push(payload.requests.map((r) => r.url));
        return {
          responses: payload.requests.map((r) => {
            const [status, body] = table[r.url] ?? [404, { error: { message: "not found" } }];
            return { id: r.id, status, body };
          }),
        };
      },
    }),
  } as unknown as GraphClient;
  return { graph, batches };
}

// Found live: /me/planner/plans returned nothing while the user's team held a plan.
describe("listPlans", () => {
  it("adds each joined team's plans to the ones shared directly, marked with their group", async () => {
    const { graph, batches } = batchGraph({
      "/me/planner/plans": [200, { value: [{ id: "shared", title: "Shared with me" }] }],
      "/me/joinedTeams?$select=id,displayName": [200, { value: [{ id: "g1", displayName: "Team One" }, { id: "g2", displayName: "Team Two" }] }],
      "/groups/g1/planner/plans": [200, { value: [{ id: "p1", title: "Roadmap" }, { id: "shared", title: "Shared with me" }] }],
      "/groups/g2/planner/plans": [403, { error: { message: "Forbidden" } }],
    });
    const out = await listPlans(graph);
    expect(out.value.map((p) => p.id).sort()).toEqual(["p1", "shared"]);
    expect(out.value.find((p) => p.id === "p1")?.group).toEqual({ id: "g1", displayName: "Team One" });
    expect(out.unreadable).toEqual([{ group: "Team Two", error: "Forbidden" }]);
    expect(batches).toHaveLength(2);
  });

  it("lists one group's plans when asked", async () => {
    const { graph } = batchGraph({ "/groups/g9/planner/plans": [200, { value: [{ id: "p9" }] }] });
    expect((await listPlans(graph, "g9")).value).toEqual([{ id: "p9", group: { id: "g9" } }]);
  });
});
