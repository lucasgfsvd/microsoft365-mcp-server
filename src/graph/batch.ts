import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";

/** Graph rejects a $batch payload carrying more than 20 requests. */
export const MAX_BATCH_SIZE = 20;

export interface BatchItem {
  /** Caller-chosen label, echoed back so responses can be matched up. */
  id: string;
  /** Graph-relative URL, e.g. "/me/messages?$top=5". */
  url: string;
}

export interface BatchItemResult {
  id: string;
  url: string;
  status: number;
  /** Present when status is 2xx. */
  body?: unknown;
  /** Present otherwise — Graph's own error, per sub-request. */
  error?: { code?: string; message?: string };
}

interface RawBatchResponse {
  responses?: Array<{
    id: string;
    status: number;
    body?: { error?: { code?: string; message?: string } } & Record<string, unknown>;
  }>;
}

/**
 * Reject anything that is not a plain Graph-relative path.
 *
 * This is the boundary that keeps batching from becoming a way to reach
 * arbitrary hosts: a sub-request URL is spliced into the batch payload and
 * resolved by Graph, so an absolute URL would send the caller's token
 * somewhere we never intended.
 */
export function assertBatchableUrl(url: string): void {
  if (!url.startsWith("/")) {
    throw new Error(`Batch url must be Graph-relative and start with "/": ${url}`);
  }
  if (url.startsWith("//")) {
    throw new Error(`Batch url must not be protocol-relative: ${url}`);
  }
  if (/^\/+[a-z][a-z0-9+.-]*:/i.test(url) || /^https?:/i.test(url)) {
    throw new Error(`Batch url must not contain a scheme: ${url}`);
  }
  if (url.split(/[/?#]/).includes("..")) {
    throw new Error(`Batch url must not contain ".." segments: ${url}`);
  }
  for (let i = 0; i < url.length; i++) {
    const c = url.charCodeAt(i);
    if (c < 0x20 || c === 0x7f) {
      throw new Error("Batch url must not contain control characters");
    }
  }
}

/**
 * Issue several Graph GETs as one $batch round trip.
 *
 * Reads only. Writes stay on their own tools so `writeGuard` remains the single
 * place mutations are authorised — batching must not become a side door around it.
 *
 * A failing sub-request does not fail the batch: each result carries its own
 * status, so a caller can use the parts that succeeded.
 */
export async function batchGet(graph: GraphClient, items: BatchItem[]): Promise<BatchItemResult[]> {
  if (items.length === 0) return [];
  if (items.length > MAX_BATCH_SIZE) {
    throw new Error(`A batch takes at most ${MAX_BATCH_SIZE} requests; got ${items.length}.`);
  }
  const seen = new Set<string>();
  for (const it of items) {
    if (seen.has(it.id)) throw new Error(`Duplicate batch id: ${it.id}`);
    seen.add(it.id);
    assertBatchableUrl(it.url);
  }

  const payload = { requests: items.map((it) => ({ id: it.id, method: "GET", url: it.url })) };
  const res = (await graph.api("/$batch").post(payload)) as RawBatchResponse;

  const byId = new Map((res.responses ?? []).map((r) => [r.id, r]));
  return items.map(({ id, url }) => {
    const r = byId.get(id);
    if (!r) {
      return { id, url, status: 0, error: { code: "NoResponse", message: "Graph returned no response for this id." } };
    }
    if (r.status >= 200 && r.status < 300) return { id, url, status: r.status, body: r.body };
    return {
      id,
      url,
      status: r.status,
      error: {
        code: r.body?.error?.code ?? `HTTP${r.status}`,
        message: r.body?.error?.message ?? "Sub-request failed.",
      },
    };
  });
}
