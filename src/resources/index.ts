import type { Server } from "@modelcontextprotocol/sdk/server/index.js";
import {
  ListResourcesRequestSchema,
  ListResourceTemplatesRequestSchema,
  ReadResourceRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { batchGet } from "../graph/batch.js";
import { logger } from "../util/logger.js";
import { driveResource } from "./drive.js";
import { mailResource } from "./mail.js";
import { onenoteResource } from "./onenote.js";
import { parseUri, type ResourceEntry, type ResourceKind } from "./types.js";

export const RESOURCE_KINDS: ResourceKind[] = [mailResource, driveResource, onenoteResource];

/** Kinds whose read tool is visible: a disabled tool is not reachable by the back door. */
export function availableKinds(tools: ReadonlySet<string>): ResourceKind[] {
  return RESOURCE_KINDS.filter((k) => tools.has(k.requiresTool));
}

/** Recent items of every available kind, fetched in one $batch. A failing kind is left out, not fatal. */
export async function listRecent(graph: GraphClient, kinds: ResourceKind[]): Promise<ResourceEntry[]> {
  if (!kinds.length) return [];
  const results = await batchGet(graph, kinds.map((k) => k.recent));
  return kinds.flatMap((k) => {
    const r = results.find((x) => x.id === k.recent.id);
    if (!r || r.status >= 300) {
      // An account with no notebook, for one, answers OneNote listings with an error.
      logger.debug({ kind: k.key, status: r?.status, error: r?.error }, "resources: recent listing unavailable");
      return [];
    }
    return k.toEntries(r.body);
  });
}

export interface ResourceDeps {
  graph: GraphClient;
  visibleTools: () => ReadonlySet<string>;
  /** Whether a token can be spent silently; never prompts. */
  signedIn: () => Promise<boolean>;
  signInHint: string;
}

export function registerResources(server: Server, deps: ResourceDeps): void {
  server.setRequestHandler(ListResourcesRequestSchema, async () => {
    // Signed out: an empty list, not an error, so the client's picker still opens.
    if (!(await deps.signedIn())) return { resources: [] };
    return { resources: await listRecent(deps.graph, availableKinds(deps.visibleTools())) };
  });

  server.setRequestHandler(ListResourceTemplatesRequestSchema, async () => ({
    resourceTemplates: availableKinds(deps.visibleTools()).map((k) => k.template),
  }));

  server.setRequestHandler(ReadResourceRequestSchema, async (req) => {
    const uri = req.params.uri;
    const parsed = parseUri(uri);
    const kind = parsed && availableKinds(deps.visibleTools()).find((k) => k.key === parsed.key);
    if (!parsed || !kind) throw new Error(`Unknown or unavailable resource: ${uri}`);
    if (!(await deps.signedIn())) throw new Error(deps.signInHint);
    return { contents: [await kind.read(deps.graph, parsed.segments, uri)] };
  });
}
