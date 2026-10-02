import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { batchGet, MAX_BATCH_SIZE } from "./batch.js";

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * Directory ids for people given by email or id, in the order given. Emails are
 * looked up in one batch; ids pass straight through. Anyone not found is an
 * error, not a silent omission: a task assigned to nobody looks like success.
 */
export async function resolveUserIds(graph: GraphClient, refs: string[]): Promise<string[]> {
  const lookups = [...new Set(refs.filter((r) => !GUID.test(r)))];
  if (!lookups.length) return refs;
  if (lookups.length > MAX_BATCH_SIZE) throw new Error(`At most ${MAX_BATCH_SIZE} people by email at a time.`);
  const found = await batchGet(graph, lookups.map((u, i) => ({ id: String(i), url: `/users/${encodeURIComponent(u)}?$select=id` })));
  const ids = new Map<string, string>();
  found.forEach((r, i) => {
    const id = (r.body as { id?: string } | undefined)?.id;
    if (!id) throw new Error(`No one with the address ${lookups[i]} in the directory: ${r.error?.message ?? "not found"}`);
    ids.set(lookups[i]!, id);
  });
  return refs.map((r) => ids.get(r) ?? r);
}
