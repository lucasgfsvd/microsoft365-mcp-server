import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";

/**
 * Point `/me` requests at a named user instead.
 *
 * Tools address the signed-in user as `/me`. That has no meaning for an
 * app-only token, which has no user (Graph refuses it), and the wrong meaning
 * for a shared mailbox. Graph's own equivalent is `/users/{id or UPN}`, which
 * accepts every path `/me` does, so one rewrite here serves every tool, and the
 * $batch sub-requests the prompts and resources build.
 */

const TARGET = Symbol("targetUser");

export interface UserTarget {
  /** UPN, address or object id; undefined means the signed-in user. */
  user?: string;
  /** No signed-in user exists (app-only), so `/me` cannot fall back to one. */
  required: boolean;
}

export const NO_USER_MESSAGE =
  "App-only mode has no signed-in user, so this tool needs one named: set MCP_USER to the user " +
  "(or shared mailbox) whose mail, calendar and files the server works with.";

const isMe = (url: string) => /^\/me(?=$|[/?])/.test(url);

export function retarget(url: string, target: UserTarget): string {
  if (!isMe(url)) return url;
  if (!target.user) {
    if (target.required) throw new Error(NO_USER_MESSAGE);
    return url;
  }
  return `/users/${encodeURIComponent(target.user)}${url.slice(3)}`;
}

/** The client as seen by requests made for `target`. Anything not under /me is untouched. */
export function graphForUser(graph: GraphClient, target: UserTarget): GraphClient {
  if (!target.user && !target.required) return graph;
  return new Proxy(graph, {
    get(inner, prop, receiver) {
      if (prop === TARGET) return target;
      if (prop === "api") return (path: string) => inner.api(retarget(path, target));
      const value = Reflect.get(inner, prop, receiver);
      return typeof value === "function" ? value.bind(inner) : value;
    },
  });
}

/** Rewrites a URL the way `graph` would; for URLs that travel inside a request body ($batch). */
export function urlRewriter(graph: GraphClient): (url: string) => string {
  const target = (graph as unknown as Record<symbol, UserTarget | undefined>)[TARGET];
  return target ? (url) => retarget(url, target) : (url) => url;
}
