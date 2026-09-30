import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { batchGet, MAX_BATCH_SIZE, type BatchItemResult } from "./batch.js";

interface Plan {
  id: string;
  title?: string;
  [key: string]: unknown;
}

export interface PlanListing {
  value: Array<Plan & { group?: { id: string; displayName?: string } }>;
  /** Teams whose plans could not be read, so the caller knows the list may be short. */
  unreadable?: Array<{ group: string; error?: string }>;
}

const plansOf = (r: BatchItemResult | undefined): Plan[] =>
  r && r.status < 300 ? ((r.body as { value?: Plan[] } | undefined)?.value ?? []) : [];

/**
 * The user's Planner plans.
 *
 * `/me/planner/plans` covers only plans shared with the user directly; plans
 * owned by their groups (the usual kind, one per team) are not in it. Seen live:
 * it returned nothing while the user's team held a plan. So this also asks each
 * joined team's group, batched. Enumerating *every* group the user belongs to
 * needs an admin-consented directory permission; teams plus direct shares is
 * what delegated Tasks.Read and Team.ReadBasic.All reach.
 */
export async function listPlans(graph: GraphClient, groupId?: string): Promise<PlanListing> {
  if (groupId) {
    const [r] = await batchGet(graph, [{ id: "group", url: `/groups/${encodeURIComponent(groupId)}/planner/plans` }]);
    if (!r || r.status >= 300) throw new Error(`Could not read plans for group ${groupId}: ${r?.error?.message ?? "no response"}`);
    return { value: plansOf(r).map((p) => ({ ...p, group: { id: groupId } })) };
  }

  const [mine, teamsRes] = await batchGet(graph, [
    { id: "me", url: "/me/planner/plans" },
    { id: "teams", url: "/me/joinedTeams?$select=id,displayName" },
  ]);
  const teams = ((teamsRes?.body as { value?: Array<{ id: string; displayName?: string }> } | undefined)?.value ?? []);

  const byId = new Map<string, PlanListing["value"][number]>();
  for (const p of plansOf(mine)) byId.set(p.id, p);

  const unreadable: NonNullable<PlanListing["unreadable"]> = [];
  for (let i = 0; i < teams.length; i += MAX_BATCH_SIZE) {
    const chunk = teams.slice(i, i + MAX_BATCH_SIZE);
    const results = await batchGet(graph, chunk.map((t) => ({ id: t.id, url: `/groups/${encodeURIComponent(t.id)}/planner/plans` })));
    for (const t of chunk) {
      const r = results.find((x) => x.id === t.id);
      if (!r || r.status >= 300) {
        unreadable.push({ group: t.displayName ?? t.id, error: r?.error?.message });
        continue;
      }
      for (const p of plansOf(r)) byId.set(p.id, { ...p, group: { id: t.id, displayName: t.displayName } });
    }
  }
  return unreadable.length ? { value: [...byId.values()], unreadable } : { value: [...byId.values()] };
}
