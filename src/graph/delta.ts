import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";

export const DELTA_RESOURCES = ["mail", "calendar", "drive", "contacts", "todo"] as const;
export type DeltaResource = (typeof DELTA_RESOURCES)[number];

export interface DeltaParams {
  /** Mail folder id or well-known name. Defaults to the inbox. */
  folderId?: string;
  /** To Do list id. Required for `todo`. */
  listId?: string;
  /** Window bounds (ISO 8601). Required for `calendar`. */
  startDateTime?: string;
  endDateTime?: string;
  select?: string[];
}

interface Descriptor {
  path: (p: DeltaParams) => string;
  query?: (p: DeltaParams) => Record<string, string>;
  /** Whether the endpoint honours $select. calendarView/delta and To Do do not. */
  selectable: boolean;
  /** Applied when the caller gives no select, to keep bodies out of the result. */
  defaultSelect?: string[];
  /** Drive ignores `Prefer: odata.maxpagesize` but honours $top. */
  pageSizeViaTop?: boolean;
}

const enc = encodeURIComponent;

/**
 * There is no single delta endpoint: each resource has its own path and its own
 * required inputs. Everything resource-specific lives here, so fetchDelta stays
 * generic.
 */
const DESCRIPTORS: Record<DeltaResource, Descriptor> = {
  // There is no /me/messages/delta; delta is per folder.
  mail: {
    path: (p) => `/me/mailFolders/${enc(p.folderId ?? "inbox")}/messages/delta`,
    selectable: true,
    defaultSelect: ["id", "subject", "from", "receivedDateTime", "isRead", "bodyPreview", "parentFolderId"],
  },
  // calendarView needs its window on the initial request; the deltaLink carries it after that.
  calendar: {
    path: () => "/me/calendarView/delta",
    query: (p) => {
      if (!p.startDateTime || !p.endDateTime) {
        throw new Error("calendar delta needs startDateTime and endDateTime on the initial request.");
      }
      return { startDateTime: p.startDateTime, endDateTime: p.endDateTime };
    },
    selectable: false,
  },
  drive: { path: () => "/me/drive/root/delta", selectable: true, pageSizeViaTop: true },
  contacts: { path: () => "/me/contacts/delta", selectable: true },
  todo: {
    path: (p) => {
      if (!p.listId) throw new Error("todo delta needs a listId (see todo_list_lists).");
      return `/me/todo/lists/${enc(p.listId)}/tasks/delta`;
    },
    selectable: false,
  },
};

export function initialDeltaUrl(
  resource: DeltaResource,
  params: DeltaParams = {},
  pageSize?: number,
): string {
  const d = DESCRIPTORS[resource];
  const query = new URLSearchParams(d.query?.(params));
  if (pageSize && d.pageSizeViaTop) query.set("$top", String(pageSize));
  const select = params.select?.length ? params.select : d.defaultSelect;
  if (select?.length) {
    if (!d.selectable) throw new Error(`${resource} delta does not support select.`);
    query.set("$select", select.join(","));
  }
  const qs = query.toString();
  return qs ? `${d.path(params)}?${qs}` : d.path(params);
}

/**
 * A cursor is handed back by us and returned by the caller, so treat it as
 * untrusted: the Graph client attaches the bearer token to whatever URL it is
 * given, and an arbitrary host here would receive it.
 */
export function assertGraphCursor(cursor: string): void {
  let url: URL;
  try {
    url = new URL(cursor);
  } catch {
    throw new Error("cursor must be the deltaLink or nextLink returned by a previous graph_delta call.");
  }
  if (url.protocol !== "https:" || url.hostname !== "graph.microsoft.com" || !/\/delta\b/.test(url.pathname)) {
    throw new Error("cursor must be a https://graph.microsoft.com delta URL from a previous graph_delta call.");
  }
}

export interface DeltaRemoval {
  id: string;
  reason: string;
}

export interface DeltaOutcome {
  /** Items added or modified since the cursor. */
  changed: Record<string, unknown>[];
  /** Items deleted (or moved out of scope) since the cursor. */
  removed: DeltaRemoval[];
  /** true when the sync reached the end; store deltaLink for next time. */
  complete: boolean;
  /** Present when complete: pass back as `cursor` to get only later changes. */
  deltaLink?: string;
  /** Present when not complete: pass back as `cursor` to continue this sync. */
  nextLink?: string;
  pages: number;
}

interface RawDeltaPage {
  value?: Array<Record<string, unknown>>;
  "@odata.nextLink"?: string;
  "@odata.deltaLink"?: string;
}

/**
 * A removal is signalled two ways: `@removed` on Outlook-family resources, and a
 * `deleted` facet on drive items. Either way the caller must see it — being told
 * what disappeared is the reason to use delta over a listing at all.
 */
function removalOf(item: Record<string, unknown>): DeltaRemoval | undefined {
  const removed = item["@removed"] as { reason?: string } | undefined;
  if (removed) return { id: String(item.id), reason: removed.reason ?? "deleted" };
  const deleted = item.deleted as { state?: string } | undefined;
  if (deleted) return { id: String(item.id), reason: deleted.state ?? "deleted" };
  return undefined;
}

function stripAnnotations(item: Record<string, unknown>): Record<string, unknown> {
  return Object.fromEntries(Object.entries(item).filter(([k]) => !k.startsWith("@odata.")));
}

/**
 * Fetch changes, following nextLinks until Graph issues a deltaLink.
 *
 * Stops early after `maxItems`, returning the nextLink instead so a first sync
 * of a large mailbox cannot blow up one tool result — `complete: false` tells
 * the caller the picture is partial. The server keeps no sync state: the link
 * goes back to the caller, who decides whether and where to keep it.
 */
export async function fetchDelta(
  graph: GraphClient,
  start: { cursor: string } | { resource: DeltaResource; params?: DeltaParams },
  opts: { maxItems?: number } = {},
): Promise<DeltaOutcome> {
  const maxItems = opts.maxItems ?? 500;
  let url: string;
  if ("cursor" in start) {
    assertGraphCursor(start.cursor);
    url = start.cursor;
  } else {
    url = initialDeltaUrl(start.resource, start.params, maxItems);
  }

  const out: DeltaOutcome = { changed: [], removed: [], complete: false, pages: 0 };
  for (;;) {
    // Graph picks the page size unless asked (10 messages, ~200 drive items), so
    // without this maxItems could only cut between pages. Outlook resources honour
    // the header; drive takes $top on the initial URL instead, and carries it on.
    const remaining = maxItems - out.changed.length - out.removed.length;
    const page = (await graph
      .api(url)
      .header("Prefer", `odata.maxpagesize=${Math.max(1, remaining)}`)
      .get()) as RawDeltaPage;
    out.pages++;
    for (const item of page.value ?? []) {
      const removal = removalOf(item);
      if (removal) out.removed.push(removal);
      else out.changed.push(stripAnnotations(item));
    }

    if (page["@odata.deltaLink"]) {
      out.complete = true;
      out.deltaLink = page["@odata.deltaLink"];
      return out;
    }
    const next = page["@odata.nextLink"];
    if (!next) {
      // Neither link: Graph broke its contract. Say so rather than pretend completeness.
      throw new Error("Graph returned a delta page with neither a nextLink nor a deltaLink.");
    }
    if (out.changed.length + out.removed.length >= maxItems) {
      out.nextLink = next;
      return out;
    }
    url = next;
  }
}
