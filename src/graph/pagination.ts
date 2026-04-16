import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";

export interface PagedResult<T> {
  value: T[];
  nextLink?: string;
}

export async function fetchPage<T>(
  graph: GraphClient,
  path: string,
  opts: { top?: number; skip?: number; filter?: string; search?: string; orderBy?: string; select?: string[]; nextLink?: string } = {},
): Promise<PagedResult<T>> {
  if (opts.nextLink) {
    const res = (await graph.api(opts.nextLink).get()) as { value: T[]; "@odata.nextLink"?: string };
    return { value: res.value, nextLink: res["@odata.nextLink"] };
  }

  let req = graph.api(path);
  if (opts.top) req = req.top(opts.top);
  if (opts.skip) req = req.skip(opts.skip);
  if (opts.filter) req = req.filter(opts.filter);
  if (opts.search) req = req.search(`"${opts.search.replace(/"/g, '\\"')}"`);
  if (opts.orderBy) req = req.orderby(opts.orderBy);
  if (opts.select && opts.select.length) req = req.select(opts.select.join(","));

  const res = (await req.get()) as { value: T[]; "@odata.nextLink"?: string };
  return { value: res.value, nextLink: res["@odata.nextLink"] };
}
