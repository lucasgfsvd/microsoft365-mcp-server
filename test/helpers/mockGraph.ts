import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import type { TokenCredential } from "@azure/identity";
import type { ServerConfig, ToolContext, ToolDefinition } from "../../src/types.js";
import type { AuthSession } from "../../src/auth/session.js";

// ---------------------------------------------------------------------------
// Test harness — a fake Graph client that records calls instead of hitting the
// network. We test that each tool constructs the right URL + payload; whether
// Graph's HTTP layer works is Microsoft's problem.
// ---------------------------------------------------------------------------

export type GraphCall = {
  path: string;
  method: "GET" | "POST" | "PATCH" | "PUT" | "DELETE";
  body?: unknown;
  query: Record<string, unknown>;
  headers: Record<string, string>;
};

type ResponseSpec = unknown | ((call: GraphCall) => unknown);

export interface MockGraph {
  graph: GraphClient;
  calls: GraphCall[];
  /** Set a response for the next call matching `path` (literal substring). */
  on(pathFragment: string, response: ResponseSpec): void;
}

export function makeMockGraph(): MockGraph {
  const calls: GraphCall[] = [];
  const responses: Array<{ pathFragment: string; response: ResponseSpec }> = [];

  function pickResponse(path: string, defaultValue: unknown, call: GraphCall): unknown {
    const idx = responses.findIndex((r) => path.includes(r.pathFragment));
    if (idx >= 0) {
      const r = responses[idx]!;
      responses.splice(idx, 1);
      return typeof r.response === "function" ? (r.response as (c: GraphCall) => unknown)(call) : r.response;
    }
    return defaultValue;
  }

  function makeBuilder(path: string): unknown {
    const query: Record<string, unknown> = {};
    const headers: Record<string, string> = {};

    const builder: Record<string, unknown> = {
      get: () => {
        const call: GraphCall = { path, method: "GET", query: { ...query }, headers: { ...headers } };
        calls.push(call);
        return Promise.resolve(pickResponse(path, { value: [] }, call));
      },
      post: (body: unknown) => {
        const call: GraphCall = { path, method: "POST", body, query: { ...query }, headers: { ...headers } };
        calls.push(call);
        return Promise.resolve(pickResponse(path, { id: "new-id" }, call));
      },
      patch: (body: unknown) => {
        const call: GraphCall = { path, method: "PATCH", body, query: { ...query }, headers: { ...headers } };
        calls.push(call);
        return Promise.resolve(pickResponse(path, {}, call));
      },
      put: (body: unknown) => {
        const call: GraphCall = { path, method: "PUT", body, query: { ...query }, headers: { ...headers } };
        calls.push(call);
        return Promise.resolve(pickResponse(path, {}, call));
      },
      delete: () => {
        const call: GraphCall = { path, method: "DELETE", query: { ...query }, headers: { ...headers } };
        calls.push(call);
        return Promise.resolve(undefined);
      },
      header: (k: string, v: string) => {
        headers[k] = v;
        return builder;
      },
      top: (n: number) => {
        query.top = n;
        return builder;
      },
      skip: (n: number) => {
        query.skip = n;
        return builder;
      },
      filter: (s: string) => {
        query.filter = s;
        return builder;
      },
      search: (s: string) => {
        query.search = s;
        return builder;
      },
      orderby: (s: string) => {
        query.orderby = s;
        return builder;
      },
      select: (s: string) => {
        query.select = s;
        return builder;
      },
      getStream: () => Promise.reject(new Error("getStream not mocked")),
    };
    return builder;
  }

  const graph = { api: makeBuilder } as unknown as GraphClient;

  return {
    graph,
    calls,
    on(pathFragment: string, response: ResponseSpec): void {
      responses.push({ pathFragment, response });
    },
  };
}

const fakeCredential: TokenCredential = {
  getToken: async () => ({ token: "fake-token", expiresOnTimestamp: Date.now() + 3_600_000 }),
};

export function makeContext(overrides: Partial<ServerConfig> = {}): ToolContext & { mock: MockGraph } {
  const mock = makeMockGraph();
  const config: ServerConfig = {
    authMode: "device-code",
    tenantId: "common",
    clientId: "x",
    tokenCachePath: "/tmp/tc.json",
    scopes: [],
    enableWrites: true,
    perSurfaceWrites: {},
    disabledTools: new Set(),
    logLevel: "info",
    listTools: false,
    logout: false,
    ...overrides,
  };
  // Tool handlers never reach for the session directly; the dispatcher gates on it.
  // A permissive stub keeps these tests about Graph calls, not sign-in.
  const auth = {
    probe: async () => true,
    status: async () => ({ state: "signed-in", authMode: config.authMode, signedIn: true }),
    signIn: async () => ({ status: "already-signed-in", message: "stub" }),
  } as unknown as AuthSession;
  return { graph: mock.graph, credential: fakeCredential, config, auth, mock };
}

export function findTool(tools: ToolDefinition[], name: string): ToolDefinition {
  const t = tools.find((x) => x.name === name);
  if (!t) throw new Error(`Tool ${name} not found`);
  return t;
}

export async function callTool(
  tool: ToolDefinition,
  ctx: ToolContext,
  rawInput: Record<string, unknown>,
): Promise<unknown> {
  const parsed = tool.inputSchema.parse(rawInput);
  return tool.handler(parsed, ctx);
}
