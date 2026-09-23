import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { createHash } from "node:crypto";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { Readable } from "node:stream";
import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { downloadToDir, safeFileName, toInline, INLINE_LIMIT } from "../src/graph/download.js";
import { filesDownloadTools } from "../src/tools/files/download.js";
import type { ToolContext } from "../src/types.js";

/** Graph stub: metadata from `meta`, content streamed in chunks, optionally failing midway. */
function graphWith(content: Buffer, opts: { meta?: object; failAfter?: number } = {}) {
  return {
    api: () => ({
      select: () => ({ get: async () => opts.meta }),
      getStream: async () =>
        Readable.from(
          (async function* () {
            for (let i = 0; i < content.length; i += 1024) {
              if (opts.failAfter !== undefined && i >= opts.failAfter) throw new Error("connection reset");
              yield content.subarray(i, i + 1024);
            }
          })(),
        ),
    }),
  } as unknown as GraphClient;
}

describe("safeFileName", () => {
  it("strips anything that could escape the folder or that Windows refuses", () => {
    expect(safeFileName("../../etc/passwd")).toBe(".._.._etc_passwd");
    expect(safeFileName("..\\..\\Windows\\win.ini")).toBe(".._.._Windows_win.ini");
    expect(safeFileName('a:b*c?"d<e>f|g')).toBe("a_b_c__d_e_f_g");
    expect(safeFileName("CON.txt")).toBe("_CON.txt");
    expect(safeFileName("trailing. . ")).toBe("trailing");
    expect(safeFileName("..")).toBe("download");
  });
});

describe("toInline", () => {
  it("returns text files as text, saving the base64 overhead", () => {
    expect(toInline(Buffer.from("héllo"), "text/plain")).toMatchObject({ encoding: "utf8", text: "héllo" });
    expect(toInline(Buffer.from("{}"), "application/json")).toMatchObject({ encoding: "utf8" });
  });

  it("falls back to base64 for binaries and for 'text' that is not UTF-8", () => {
    expect(toInline(Buffer.from([0xff, 0xfe]), "text/plain")).toMatchObject({ encoding: "base64", base64: "//4=" });
    expect(toInline(Buffer.from("x"), "application/pdf")).toMatchObject({ encoding: "base64" });
  });
});

describe("downloadToDir", () => {
  let dir: string;
  beforeEach(async () => {
    dir = await fs.mkdtemp(path.join(os.tmpdir(), "m365-dl-"));
  });
  afterEach(async () => {
    await fs.rm(dir, { recursive: true, force: true });
  });

  it("streams to disk and reports size and checksum", async () => {
    const data = Buffer.alloc(10_000, 7);
    const saved = await downloadToDir(graphWith(data), "/c", dir, "report.pdf");
    expect(saved).toEqual({
      path: path.join(dir, "report.pdf"),
      byteLength: data.length,
      sha256: createHash("sha256").update(data).digest("hex"),
    });
    expect(await fs.readFile(saved.path)).toEqual(data);
  });

  it("never overwrites: a taken name gets a numbered suffix", async () => {
    await fs.writeFile(path.join(dir, "report.pdf"), "precious");
    const saved = await downloadToDir(graphWith(Buffer.from("new")), "/c", dir, "report.pdf");
    expect(path.basename(saved.path)).toBe("report (1).pdf");
    expect(await fs.readFile(path.join(dir, "report.pdf"), "utf8")).toBe("precious");
  });

  it("keeps a hostile name inside the folder", async () => {
    const saved = await downloadToDir(graphWith(Buffer.from("x")), "/c", dir, "../../escape.txt");
    expect(path.dirname(saved.path)).toBe(path.resolve(dir));
  });

  it("removes the partial file when the transfer fails", async () => {
    await expect(
      downloadToDir(graphWith(Buffer.alloc(8192), { failAfter: 4096 }), "/c", dir, "big.bin"),
    ).rejects.toThrow(/connection reset/);
    expect(await fs.readdir(dir)).toEqual([]);
  });
});

describe("files_download", () => {
  const tool = filesDownloadTools[0];
  const ctx = (graph: GraphClient, downloadDir?: string) => ({ graph, config: { downloadDir } }) as unknown as ToolContext;

  it("refuses an oversized inline download before fetching it, and says how to get it", async () => {
    const graph = graphWith(Buffer.alloc(0), { meta: { name: "huge.zip", size: INLINE_LIMIT + 1, file: {} } });
    await expect(tool.handler({ itemId: "i" }, ctx(graph))).rejects.toThrow(/MCP_DOWNLOAD_DIR.*saveToDisk/);
  });

  it("will not save to disk unless a folder is configured", async () => {
    const graph = graphWith(Buffer.from("x"), { meta: { name: "a.txt", size: 1, file: {} } });
    await expect(tool.handler({ itemId: "i", saveToDisk: true }, ctx(graph))).rejects.toThrow(/MCP_DOWNLOAD_DIR/);
  });

  it("returns small text files inline as text", async () => {
    const graph = graphWith(Buffer.from("hello"), { meta: { name: "a.txt", size: 5, file: { mimeType: "text/plain" } } });
    expect(await tool.handler({ itemId: "i" }, ctx(graph))).toMatchObject({ name: "a.txt", encoding: "utf8", text: "hello" });
  });
});
