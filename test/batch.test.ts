import { describe, it, expect } from "vitest";
import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { assertBatchableUrl, batchGet, MAX_BATCH_SIZE } from "../src/graph/batch.js";

/** Minimal fake: records the posted payload, replies with a canned response. */
function mockGraph(reply: unknown) {
  const posted: unknown[] = [];
  const graph = {
    api: (path: string) => ({
      post: (body: unknown) => {
        posted.push({ path, body });
        return Promise.resolve(reply);
      },
    }),
  } as unknown as GraphClient;
  return { graph, posted };
}

describe("assertBatchableUrl", () => {
  it("accepts Graph-relative paths, including query strings", () => {
    for (const url of ["/me/messages", "/me/messages?$top=5&$select=subject", "/me/drive/root:/a b.txt:/content"]) {
      expect(() => assertBatchableUrl(url)).not.toThrow();
    }
  });

  // This is the boundary that stops a sub-request sending the token elsewhere.
  it("rejects absolute and protocol-relative urls", () => {
    for (const url of ["https://evil.example/steal", "http://evil.example", "//evil.example/x"]) {
      expect(() => assertBatchableUrl(url)).toThrow();
    }
  });

  it("rejects a scheme smuggled behind a leading slash", () => {
    expect(() => assertBatchableUrl("/https://evil.example")).toThrow(/scheme/);
  });

  it("rejects paths that are not rooted", () => {
    expect(() => assertBatchableUrl("me/messages")).toThrow(/start with/);
  });

  it("rejects upward traversal", () => {
    expect(() => assertBatchableUrl("/me/../../admin")).toThrow(/\.\./);
  });

  it("rejects control characters", () => {
    expect(() => assertBatchableUrl("/me/mess\nages")).toThrow(/control/);
  });
});

describe("batchGet", () => {
  it("posts one $batch with GET sub-requests", async () => {
    const { graph, posted } = mockGraph({ responses: [{ id: "a", status: 200, body: { value: [1] } }] });
    await batchGet(graph, [{ id: "a", url: "/me/messages" }]);
    expect(posted).toHaveLength(1);
    expect(posted[0]).toEqual({
      path: "/$batch",
      body: { requests: [{ id: "a", method: "GET", url: "/me/messages" }] },
    });
  });

  it("returns results in the order requested, not the order replied", async () => {
    const { graph } = mockGraph({
      responses: [
        { id: "b", status: 200, body: { v: 2 } },
        { id: "a", status: 200, body: { v: 1 } },
      ],
    });
    const out = await batchGet(graph, [
      { id: "a", url: "/me/a" },
      { id: "b", url: "/me/b" },
    ]);
    expect(out.map((r) => r.id)).toEqual(["a", "b"]);
    expect(out[0]!.body).toEqual({ v: 1 });
  });

  it("surfaces a failing sub-request without failing the batch", async () => {
    const { graph } = mockGraph({
      responses: [
        { id: "ok", status: 200, body: { value: [] } },
        { id: "bad", status: 404, body: { error: { code: "itemNotFound", message: "nope" } } },
      ],
    });
    const out = await batchGet(graph, [
      { id: "ok", url: "/me/ok" },
      { id: "bad", url: "/me/bad" },
    ]);
    expect(out[0]!.error).toBeUndefined();
    expect(out[1]!.status).toBe(404);
    expect(out[1]!.error).toEqual({ code: "itemNotFound", message: "nope" });
  });

  it("marks ids Graph did not answer", async () => {
    const { graph } = mockGraph({ responses: [] });
    const out = await batchGet(graph, [{ id: "a", url: "/me/a" }]);
    expect(out[0]!.status).toBe(0);
    expect(out[0]!.error?.code).toBe("NoResponse");
  });

  it("refuses more than Graph accepts", async () => {
    const { graph } = mockGraph({ responses: [] });
    const many = Array.from({ length: MAX_BATCH_SIZE + 1 }, (_, i) => ({ id: String(i), url: "/me/x" }));
    await expect(batchGet(graph, many)).rejects.toThrow(/at most/);
  });

  it("refuses duplicate ids, which would make responses ambiguous", async () => {
    const { graph } = mockGraph({ responses: [] });
    await expect(
      batchGet(graph, [
        { id: "a", url: "/me/a" },
        { id: "a", url: "/me/b" },
      ]),
    ).rejects.toThrow(/Duplicate/);
  });

  it("validates every url before posting anything", async () => {
    const { graph, posted } = mockGraph({ responses: [] });
    await expect(
      batchGet(graph, [
        { id: "a", url: "/me/ok" },
        { id: "b", url: "https://evil.example" },
      ]),
    ).rejects.toThrow();
    expect(posted).toHaveLength(0);
  });

  it("short-circuits an empty batch", async () => {
    const { graph, posted } = mockGraph({ responses: [] });
    expect(await batchGet(graph, [])).toEqual([]);
    expect(posted).toHaveLength(0);
  });
});
