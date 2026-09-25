import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import type { BatchItem } from "../graph/batch.js";

/**
 * MCP resources are what the *user* attaches as context (an @-mention in
 * Claude Code), where tools are what the model decides to call. Each kind here
 * lists recent items and reads any item by URI, as text wherever that is
 * possible, because the content lands in the conversation.
 */

export const SCHEME = "m365://";

export interface ResourceEntry {
  uri: string;
  name: string;
  title?: string;
  description?: string;
  mimeType?: string;
}

export type ResourceContent = { uri: string; mimeType?: string } & ({ text: string } | { blob: string });

export interface ResourceKind {
  /** First URI segment after the scheme: m365://<key>/... */
  key: string;
  /** The read tool whose visibility gates this kind, so config cannot be bypassed. */
  requiresTool: string;
  template: { uriTemplate: string; name: string; title: string; description: string; mimeType?: string };
  /** One Graph GET listing recent items; all kinds are fetched in one $batch. */
  recent: BatchItem;
  toEntries(body: unknown): ResourceEntry[];
  /** `segments` are the decoded URI path parts after the key. */
  read(graph: GraphClient, segments: string[], uri: string): Promise<ResourceContent>;
}

export const enc = encodeURIComponent;

export function uriFor(key: string, ...ids: string[]): string {
  return `${SCHEME}${key}/${ids.map(enc).join("/")}`;
}

/** Split a resource URI into its kind and decoded path segments. */
export function parseUri(uri: string): { key: string; segments: string[] } | undefined {
  if (!uri.startsWith(SCHEME)) return undefined;
  const [key, ...rest] = uri.slice(SCHEME.length).split("/");
  if (!key || rest.length === 0 || rest.some((s) => s === "")) return undefined;
  try {
    return { key, segments: rest.map(decodeURIComponent) };
  } catch {
    return undefined;
  }
}

export const values = (body: unknown): Array<Record<string, unknown>> =>
  ((body as { value?: unknown[] } | undefined)?.value ?? []) as Array<Record<string, unknown>>;
