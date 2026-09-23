import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";

/** Entity types this server exposes through the Search API. */
export const SEARCH_ENTITY_TYPES = [
  "message",
  "event",
  "driveItem",
  "listItem",
  "list",
  "site",
  "chatMessage",
  "person",
] as const;

export type SearchEntityType = (typeof SEARCH_ENTITY_TYPES)[number];

/**
 * Graph will not mix entity types across these groups in one request, and the
 * error it returns when you try ("Invalid entity type combination") does not
 * say which pairing was the problem. Checking up front gives the caller
 * something actionable.
 *
 * Only the SharePoint/OneDrive family combines. message, event, chatMessage
 * and person each have to be asked for on their own — verified against a live
 * tenant, which rejected message+event despite both being "mail-ish".
 */
const GROUPS: Record<string, readonly SearchEntityType[]> = {
  "files & sites": ["driveItem", "listItem", "list", "site"],
  mail: ["message"],
  calendar: ["event"],
  "teams messages": ["chatMessage"],
  people: ["person"],
};

export function assertCombinableEntityTypes(types: readonly SearchEntityType[]): void {
  const groups = new Set<string>();
  for (const t of types) {
    const g = Object.entries(GROUPS).find(([, members]) => members.includes(t));
    if (g) groups.add(g[0]);
  }
  if (groups.size > 1) {
    const names = [...groups].sort().join(" + ");
    throw new Error(
      `Graph cannot search ${names} in one request. Split it into one call per group ` +
        `(${Object.keys(GROUPS).join(", ")}).`,
    );
  }
}

export interface SearchHit {
  entityType: string;
  id?: string;
  name?: string;
  summary?: string;
  webUrl?: string;
  lastModified?: string;
}

export interface SearchOutcome {
  total?: number;
  moreAvailable?: boolean;
  hits: SearchHit[];
}

interface RawSearchResponse {
  value?: Array<{
    hitsContainers?: Array<{
      total?: number;
      moreResultsAvailable?: boolean;
      hits?: Array<{
        summary?: string;
        resource?: Record<string, unknown> & {
          "@odata.type"?: string;
          id?: string;
          name?: string;
          subject?: string;
          displayName?: string;
          webUrl?: string;
          webLink?: string;
          lastModifiedDateTime?: string;
        };
      }>;
    }>;
  }>;
}

/** Best-effort human label for a hit, since each entity type names itself differently. */
function hitName(r: Record<string, unknown> | undefined): string | undefined {
  if (!r) return undefined;
  return (r.name ?? r.subject ?? r.displayName) as string | undefined;
}

/**
 * Query the Microsoft Search API across several surfaces at once.
 *
 * Distinct from the per-surface `*_search` tools: those hit one endpoint each
 * with its own filter syntax, whereas this is one relevance-ranked query over
 * mail, files, SharePoint or Teams together.
 */
export async function searchQuery(
  graph: GraphClient,
  opts: { query: string; entityTypes: SearchEntityType[]; size?: number; from?: number },
): Promise<SearchOutcome> {
  assertCombinableEntityTypes(opts.entityTypes);

  const payload = {
    requests: [
      {
        entityTypes: opts.entityTypes,
        query: { queryString: opts.query },
        from: opts.from ?? 0,
        size: opts.size ?? 25,
      },
    ],
  };

  const res = (await graph.api("/search/query").post(payload)) as RawSearchResponse;
  const container = res.value?.[0]?.hitsContainers?.[0];

  return {
    total: container?.total,
    moreAvailable: container?.moreResultsAvailable,
    hits: (container?.hits ?? []).map((h) => ({
      entityType: String(h.resource?.["@odata.type"] ?? "").replace(/^#microsoft\.graph\./, "") || "unknown",
      id: h.resource?.id,
      name: hitName(h.resource),
      summary: h.summary,
      webUrl: (h.resource?.webUrl ?? h.resource?.webLink) as string | undefined,
      lastModified: h.resource?.lastModifiedDateTime,
    })),
  };
}
