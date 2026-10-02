import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { Readable } from "node:stream";
import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { progressReporter, type ProgressNotification } from "../src/util/progress.js";
import { SIMPLE_UPLOAD_LIMIT, uploadContent } from "../src/graph/upload.js";
import { downloadToDir } from "../src/graph/download.js";
import { uploadOptions } from "../src/tools/files/transfer.js";

describe("progressReporter", () => {
  function recorder(intervalMs = 250) {
    const sent: ProgressNotification["params"][] = [];
    let t = 0;
    const report = progressReporter("tok", async (n) => void sent.push(n.params), { intervalMs, now: () => t })!;
    return { sent, report, advance: (ms: number) => (t += ms) };
  }

  it("sends nothing when the client did not ask for progress", () => {
    expect(progressReporter(undefined, async () => undefined)).toBeUndefined();
  });

  it("sends the first update, then at most one per interval, and always the last", () => {
    const { sent, report, advance } = recorder();
    report(10, 100, "a");
    report(20, 100);
    advance(100);
    report(30, 100);
    advance(200);
    report(40, 100);
    report(100, 100, "done");
    expect(sent).toEqual([
      { progressToken: "tok", progress: 10, total: 100, message: "a" },
      { progressToken: "tok", progress: 40, total: 100, message: undefined },
      { progressToken: "tok", progress: 100, total: 100, message: "done" },
    ]);
  });

  it("never sends progress that does not increase, as the spec requires", () => {
    const { sent, report, advance } = recorder(0);
    report(50, 100);
    advance(1);
    report(50, 100); // a retried chunk re-reports where the session stands
    report(40, 100);
    expect(sent.map((p) => p.progress)).toEqual([50]);
  });
});

const graphSession = () => ({ api: () => ({ post: async () => ({ uploadUrl: "https://x.sharepoint.com/up" }) }) }) as unknown as GraphClient;
const UNIT = 320 * 1024;

describe("uploads", () => {
  it("report each chunk the session accepts, ending at the total", async () => {
    const size = SIMPLE_UPLOAD_LIMIT + 3 * UNIT;
    let received = 0;
    const fetchFn = (async (_u: string, init: RequestInit) => {
      received += (init.body as Buffer).length;
      const done = received === size;
      return new Response(JSON.stringify(done ? { id: "f" } : { nextExpectedRanges: [`${received}-`] }), { status: done ? 201 : 202 });
    }) as unknown as typeof fetch;
    const seen: number[] = [];
    await uploadContent(graphSession(), "/me/drive", { itemId: "i" }, Buffer.alloc(size), {
      fetch: fetchFn,
      chunkSize: 8 * UNIT,
      onProgress: (sent, total) => {
        expect(total).toBe(size);
        seen.push(sent);
      },
    });
    expect(seen).toEqual([8 * UNIT, size]);
  });

  it("stop when cancelled, sending no more chunks, and cancel the session", async () => {
    const controller = new AbortController();
    const calls: string[] = [];
    const fetchFn = (async (_u: string, init: RequestInit) => {
      calls.push(init.method!);
      if (init.method === "PUT") controller.abort(); // the user cancels while the first chunk is in flight
      return new Response(JSON.stringify({ nextExpectedRanges: [`${8 * UNIT}-`] }), { status: 202 });
    }) as unknown as typeof fetch;
    await expect(
      uploadContent(graphSession(), "/me/drive", { itemId: "i" }, Buffer.alloc(SIMPLE_UPLOAD_LIMIT * 3), {
        fetch: fetchFn,
        chunkSize: 8 * UNIT,
        signal: controller.signal,
      }),
    ).rejects.toMatchObject({ name: "AbortError" });
    expect(calls).toEqual(["PUT", "DELETE"]);
  });

  it("do not start once already cancelled", async () => {
    const puts: unknown[] = [];
    const graph = { api: () => ({ put: async (b: unknown) => puts.push(b) }) } as unknown as GraphClient;
    await expect(uploadContent(graph, "/me/drive", { itemId: "i" }, Buffer.from("x"), { signal: AbortSignal.abort() })).rejects.toThrow();
    expect(puts).toEqual([]);
  });

  it("describe themselves to the client in megabytes", () => {
    const reports: unknown[][] = [];
    const opts = uploadOptions({ progress: (...a) => void reports.push(a) }, "q3.pdf");
    opts.onProgress!(1024 * 1024, 4 * 1024 * 1024);
    expect(reports).toEqual([[1024 * 1024, 4 * 1024 * 1024, "Uploading q3.pdf: 1.0 MB of 4.0 MB"]]);
    expect(uploadOptions({}, "x").onProgress).toBeUndefined();
  });
});

describe("downloads to disk", () => {
  let dir: string;
  beforeEach(async () => (dir = await fs.mkdtemp(path.join(os.tmpdir(), "m365-progress-"))));
  afterEach(async () => fs.rm(dir, { recursive: true, force: true }));

  const streaming = (size: number, onChunk?: (i: number) => void) =>
    ({
      api: () => ({
        getStream: async () =>
          Readable.from(
            (async function* () {
              for (let i = 0; i < size; i += 64 * 1024) {
                onChunk?.(i);
                await new Promise((r) => setImmediate(r));
                yield Buffer.alloc(Math.min(64 * 1024, size - i));
              }
            })(),
          ),
      }),
    }) as unknown as GraphClient;

  it("report the bytes written", async () => {
    const seen: number[] = [];
    await downloadToDir(streaming(256 * 1024), "/x/content", dir, "big.bin", { onProgress: (n) => seen.push(n) });
    expect(seen).toEqual([65536, 131072, 196608, 262144]);
  });

  it("stop when cancelled, and leave no partial file", async () => {
    const controller = new AbortController();
    const graph = streaming(1024 * 1024, (i) => i >= 128 * 1024 && controller.abort());
    await expect(downloadToDir(graph, "/x/content", dir, "big.bin", { signal: controller.signal })).rejects.toMatchObject({ name: "AbortError" });
    expect(await fs.readdir(dir)).toEqual([]);
  });
});
