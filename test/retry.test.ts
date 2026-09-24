import { describe, it, expect, beforeEach, afterEach } from "vitest";
import nock from "nock";
import type { TokenCredential } from "@azure/identity";
import { buildGraphClient } from "../src/graph/client.js";
import { graphForTool } from "../src/graph/retry.js";

// These run the SDK's real RetryHandler against HTTP stubbed by nock, so they
// check what actually goes over the wire, not a model of it.
const GRAPH = "https://graph.microsoft.com";
const credential: TokenCredential = {
  getToken: async () => ({ token: "t", expiresOnTimestamp: Date.now() + 3_600_000 }),
};
const graph = () => buildGraphClient(credential, ["User.Read"]);
// Retry-After: 0 keeps the SDK from sleeping its 3 s default between attempts.
const now = { "Retry-After": "0" };

const sender = { mutating: true };
const reader = {};

describe("per-tool retry policy", () => {
  beforeEach(() => nock.disableNetConnect());
  afterEach(() => {
    nock.cleanAll();
    nock.enableNetConnect();
  });

  // The bug this exists for: the SDK alone retries a POST on 503, so a send
  // that Graph had already carried out would go out again.
  it("baseline: without the policy the SDK re-sends a POST after a 503", async () => {
    const scope = nock(GRAPH).post("/v1.0/me/sendMail").reply(503, {}, now).post("/v1.0/me/sendMail").reply(202);
    await graph().api("/me/sendMail").post({ message: {} });
    expect(scope.isDone()).toBe(true);
  });

  it("never re-sends a mutating POST after 503 or 504", async () => {
    for (const status of [503, 504]) {
      const scope = nock(GRAPH).post("/v1.0/me/sendMail").reply(status, {}, now);
      const second = nock(GRAPH).post("/v1.0/me/sendMail").reply(202);
      await expect(graphForTool(graph(), sender).api("/me/sendMail").post({ message: {} })).rejects.toBeDefined();
      expect(scope.isDone()).toBe(true);
      expect(second.isDone()).toBe(false);
      nock.cleanAll();
    }
  });

  it("does retry a mutating POST that was throttled (429), since Graph did not act on it", async () => {
    const scope = nock(GRAPH).post("/v1.0/me/sendMail").reply(429, {}, now).post("/v1.0/me/sendMail").reply(202);
    await graphForTool(graph(), sender).api("/me/sendMail").post({ message: {} });
    expect(scope.isDone()).toBe(true);
  });

  it("keeps retrying reads, including read-only POSTs", async () => {
    const get = nock(GRAPH).get("/v1.0/me/messages").reply(503, {}, now).get("/v1.0/me/messages").reply(200, { value: [] });
    await graphForTool(graph(), reader).api("/me/messages").get();
    expect(get.isDone()).toBe(true);

    const search = nock(GRAPH).post("/v1.0/search/query").reply(504, {}, now).post("/v1.0/search/query").reply(200, { value: [] });
    await graphForTool(graph(), reader).api("/search/query").post({ requests: [] });
    expect(search.isDone()).toBe(true);
  });

  it("keeps retrying repeatable writes (PATCH, DELETE) from mutating tools", async () => {
    const patch = nock(GRAPH).patch("/v1.0/me/events/1").reply(503, {}, now).patch("/v1.0/me/events/1").reply(200, {});
    await graphForTool(graph(), sender).api("/me/events/1").patch({ subject: "x" });
    expect(patch.isDone()).toBe(true);

    const del = nock(GRAPH).delete("/v1.0/me/events/1").reply(503, {}, now).delete("/v1.0/me/events/1").reply(204);
    await graphForTool(graph(), sender).api("/me/events/1").delete();
    expect(del.isDone()).toBe(true);
  });

  it("honours a tool's own retry count", async () => {
    const scope = nock(GRAPH).get("/v1.0/me").times(2).reply(503, {}, now);
    const third = nock(GRAPH).get("/v1.0/me").reply(200, {});
    await expect(graphForTool(graph(), { retry: { maxRetries: 1 } }).api("/me").get()).rejects.toBeDefined();
    expect(scope.isDone()).toBe(true);
    expect(third.isDone()).toBe(false);
  });
});
