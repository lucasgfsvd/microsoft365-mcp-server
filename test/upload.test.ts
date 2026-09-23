import { describe, it, expect } from "vitest";
import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { uploadContent, SIMPLE_UPLOAD_LIMIT } from "../src/graph/upload.js";

const UNIT = 320 * 1024;
const UPLOAD_URL = "https://tenant.sharepoint.com/upload/session-1";

function mockGraph() {
  const calls: Array<{ method: string; path: string; body?: unknown }> = [];
  const graph = {
    api: (path: string) => ({
      put: (body: unknown) => {
        calls.push({ method: "PUT", path, body });
        return Promise.resolve({ id: "simple" });
      },
      post: (body: unknown) => {
        calls.push({ method: "POST", path, body });
        return Promise.resolve({ uploadUrl: UPLOAD_URL });
      },
    }),
  } as unknown as GraphClient;
  return { graph, calls };
}

/**
 * A fake upload session: accepts chunks in order, reports progress on GET, and
 * can be told to fail particular PUTs — optionally after accepting the bytes, as
 * when a response is lost in transit.
 */
function fakeSession(total: number, faults: Record<number, { status: number; accept?: boolean }> = {}) {
  let received = 0;
  let put = 0;
  const log: string[] = [];
  const json = (status: number, body: unknown, headers: Record<string, string> = {}) =>
    new Response(JSON.stringify(body), { status, headers });
  const fetchFn = (async (url: string, init: RequestInit) => {
    expect(url).toBe(UPLOAD_URL);
    const headers = (init.headers ?? {}) as Record<string, string>;
    // The session URL is pre-authorised; the bearer token must never go there.
    expect(Object.keys(headers).map((h) => h.toLowerCase())).not.toContain("authorization");
    if (init.method === "GET") {
      log.push("GET");
      return json(200, { nextExpectedRanges: [`${received}-`] });
    }
    if (init.method === "DELETE") {
      log.push("DELETE");
      return new Response(null, { status: 204 });
    }
    const [, start, end] = /bytes (\d+)-(\d+)\//.exec(headers["Content-Range"])!.map(Number);
    log.push(`PUT ${start}-${end}`);
    const fault = faults[put++];
    if (fault) {
      if (fault.accept && start === received) received = end + 1;
      return json(fault.status, { error: "fault" }, fault.status === 429 ? { "Retry-After": "1" } : {});
    }
    if (start !== received) return json(416, { error: "range" });
    received = end + 1;
    if (received === total) return json(201, { id: "big", size: total });
    return json(202, { nextExpectedRanges: [`${received}-`] });
  }) as unknown as typeof fetch;
  return { fetchFn, log, received: () => received };
}

const noSleep = async () => {};

describe("uploadContent", () => {
  it("uses the single PUT at or under the limit", async () => {
    const { graph, calls } = mockGraph();
    await uploadContent(graph, "/me/drive", { parentPath: "/Reports/", filename: "a.txt" }, Buffer.alloc(SIMPLE_UPLOAD_LIMIT));
    expect(calls).toEqual([{ method: "PUT", path: "/me/drive/root:/Reports/a.txt:/content", body: expect.any(Buffer) }]);
  });

  it("opens a replace-on-conflict session by path above the limit, and sends every byte once", async () => {
    const { graph, calls } = mockGraph();
    const total = SIMPLE_UPLOAD_LIMIT + 1;
    const s = fakeSession(total);
    const item = await uploadContent(graph, "/me/drive", { parentPath: "/", filename: "big.bin" }, Buffer.alloc(total), {
      fetch: s.fetchFn,
      chunkSize: 4 * UNIT,
    });
    expect(calls).toEqual([
      {
        method: "POST",
        path: "/me/drive/root:/big.bin:/createUploadSession",
        body: { item: { "@microsoft.graph.conflictBehavior": "replace" } },
      },
    ]);
    expect(item).toEqual({ id: "big", size: total });
    const c = 4 * UNIT;
    expect(s.log).toEqual([`PUT 0-${c - 1}`, `PUT ${c}-${2 * c - 1}`, `PUT ${2 * c}-${3 * c - 1}`, `PUT ${3 * c}-${total - 1}`]);
  });

  it("addresses an existing item by id", async () => {
    const { graph, calls } = mockGraph();
    const s = fakeSession(SIMPLE_UPLOAD_LIMIT + 10);
    await uploadContent(graph, "/drives/d1", { itemId: "i1" }, Buffer.alloc(SIMPLE_UPLOAD_LIMIT + 10), { fetch: s.fetchFn });
    expect(calls[0]).toMatchObject({ path: "/drives/d1/items/i1/createUploadSession", body: {} });
  });

  it("retries a throttled chunk after asking the session where it stands", async () => {
    const { graph } = mockGraph();
    const total = 14 * UNIT; // just over the simple-upload limit
    const s = fakeSession(total, { 1: { status: 429 } });
    await uploadContent(graph, "/me/drive", { itemId: "i" }, Buffer.alloc(total), { fetch: s.fetchFn, chunkSize: UNIT, sleep: noSleep });
    expect(s.log.slice(0, 5)).toEqual([`PUT 0-${UNIT - 1}`, `PUT ${UNIT}-${2 * UNIT - 1}`, "GET", `PUT ${UNIT}-${2 * UNIT - 1}`, `PUT ${2 * UNIT}-${3 * UNIT - 1}`]);
    expect(s.received()).toBe(total);
  });

  // A chunk can land even though its response is lost; resending it would be refused.
  it("skips ahead when a failed chunk had actually been received", async () => {
    const { graph } = mockGraph();
    const total = 14 * UNIT;
    const s = fakeSession(total, { 1: { status: 503, accept: true } });
    await uploadContent(graph, "/me/drive", { itemId: "i" }, Buffer.alloc(total), { fetch: s.fetchFn, chunkSize: UNIT, sleep: noSleep });
    expect(s.log.slice(0, 4)).toEqual([`PUT 0-${UNIT - 1}`, `PUT ${UNIT}-${2 * UNIT - 1}`, "GET", `PUT ${2 * UNIT}-${3 * UNIT - 1}`]);
    expect(s.received()).toBe(total);
  });

  it("gives up after repeated failures and cancels the session", async () => {
    const { graph } = mockGraph();
    const total = 2 * UNIT;
    const s = fakeSession(total, { 0: { status: 500 }, 1: { status: 500 }, 2: { status: 500 }, 3: { status: 500 } });
    await expect(
      uploadContent(graph, "/me/drive", { itemId: "i" }, Buffer.alloc(SIMPLE_UPLOAD_LIMIT + 1), { fetch: s.fetchFn, chunkSize: UNIT, sleep: noSleep }),
    ).rejects.toThrow(/after 4 attempts/);
    expect(s.log.at(-1)).toBe("DELETE");
  });

  it("does not retry a client error, and cancels the session", async () => {
    const { graph } = mockGraph();
    const s = fakeSession(SIMPLE_UPLOAD_LIMIT + 1, { 0: { status: 400 } });
    await expect(
      uploadContent(graph, "/me/drive", { itemId: "i" }, Buffer.alloc(SIMPLE_UPLOAD_LIMIT + 1), { fetch: s.fetchFn, sleep: noSleep }),
    ).rejects.toThrow(/HTTP 400/);
    expect(s.log).toEqual([`PUT 0-${SIMPLE_UPLOAD_LIMIT}`, "DELETE"]);
  });

  it("refuses a chunk size Graph would reject", async () => {
    const { graph } = mockGraph();
    await expect(
      uploadContent(graph, "/me/drive", { itemId: "i" }, Buffer.alloc(SIMPLE_UPLOAD_LIMIT + 1), { chunkSize: 1000 }),
    ).rejects.toThrow(/multiple of/);
  });
});
