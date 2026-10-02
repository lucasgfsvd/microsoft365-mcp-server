import { describe, it, expect, beforeAll, afterAll } from "vitest";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { fileSource, resolveUploadPath, uploadLocalFile } from "../src/graph/localUpload.js";
import { SIMPLE_UPLOAD_LIMIT, uploadContent } from "../src/graph/upload.js";
import { filesTools } from "../src/tools/files/index.js";
import { callTool, findTool, makeContext } from "./helpers/mockGraph.js";

let tmp: string;
let uploads: string;
let outsideFile: string;

beforeAll(async () => {
  tmp = await fs.mkdtemp(path.join(os.tmpdir(), "m365-upload-"));
  uploads = path.join(tmp, "uploads");
  await fs.mkdir(path.join(uploads, "sub"), { recursive: true });
  await fs.writeFile(path.join(uploads, "report.txt"), "inside");
  await fs.writeFile(path.join(uploads, "sub", "deep.txt"), "deeper");
  await fs.writeFile(path.join(uploads, "..dots.txt"), "a name, not a parent");
  outsideFile = path.join(tmp, "secret.txt");
  await fs.writeFile(outsideFile, "outside");
});

afterAll(async () => {
  await fs.rm(tmp, { recursive: true, force: true });
});

describe("resolveUploadPath", () => {
  it("finds files inside the folder, by relative or absolute path", async () => {
    const real = await fs.realpath(path.join(uploads, "sub", "deep.txt"));
    expect(await resolveUploadPath(uploads, "sub/deep.txt")).toBe(real);
    expect(await resolveUploadPath(uploads, path.join(uploads, "sub", "deep.txt"))).toBe(real);
    expect(await resolveUploadPath(uploads, "..dots.txt")).toMatch(/\.\.dots\.txt$/);
  });

  it("refuses anything outside, without saying whether it exists", async () => {
    for (const p of ["../secret.txt", "../nothing-here.txt", outsideFile, "sub/../../secret.txt"]) {
      await expect(resolveUploadPath(uploads, p)).rejects.toThrow(/outside the upload folder/);
    }
  });

  it("refuses a link inside the folder that leads out of it", async () => {
    const link = path.join(uploads, "escape");
    try {
      await fs.symlink(tmp, link, "junction"); // a junction needs no privileges on Windows
    } catch {
      return; // links unavailable here; nothing to test
    }
    await expect(resolveUploadPath(uploads, "escape/secret.txt")).rejects.toThrow(/outside the upload folder/);
  });

  it("refuses folders, missing files, NUL bytes and a missing upload folder", async () => {
    await expect(resolveUploadPath(uploads, "sub")).rejects.toThrow(/not a file/);
    await expect(resolveUploadPath(uploads, ".")).rejects.toThrow(/outside/);
    await expect(resolveUploadPath(uploads, "missing.txt")).rejects.toThrow(/No file at missing.txt/);
    await expect(resolveUploadPath(uploads, "a\0b")).rejects.toThrow(/NUL/);
    await expect(resolveUploadPath(path.join(tmp, "nope"), "x")).rejects.toThrow(/does not exist/);
  });
});

describe("streaming from disk", () => {
  it("reads any range of a file, short only at its end", async () => {
    const file = path.join(uploads, "report.txt");
    const handle = await fs.open(file, "r");
    try {
      const src = fileSource(handle, 6);
      expect((await src.read(1, 3)).toString()).toBe("nsi");
      expect((await src.read(4, 10)).toString()).toBe("de");
    } finally {
      await handle.close();
    }
  });

  it("sends a small file in one PUT", async () => {
    const puts: Array<{ path: string; body: Buffer }> = [];
    const graph = { api: (p: string) => ({ put: async (body: Buffer) => (puts.push({ path: p, body }), { id: "f" }) }) } as unknown as GraphClient;
    await uploadLocalFile(graph, "/me/drive", { parentPath: "/Docs", filename: "r.txt" }, uploads, "report.txt");
    expect(puts).toEqual([{ path: "/me/drive/root:/Docs/r.txt:/content", body: Buffer.from("inside") }]);
  });

  it("sends a large file through a session, chunk by chunk, byte for byte", async () => {
    const big = Buffer.alloc(SIMPLE_UPLOAD_LIMIT + 700 * 1024);
    for (let i = 0; i < big.length; i++) big[i] = (i * 7) % 251;
    await fs.writeFile(path.join(uploads, "big.bin"), big);
    const graph = { api: () => ({ post: async () => ({ uploadUrl: "https://x.sharepoint.com/up" }) }) } as unknown as GraphClient;
    const received: Buffer[] = [];
    const fetchFn = (async (_url: string, init: RequestInit) => {
      received.push(Buffer.from(init.body as Buffer));
      const sent = received.reduce((n, b) => n + b.length, 0);
      return new Response(JSON.stringify(sent === big.length ? { id: "big" } : { nextExpectedRanges: [`${sent}-`] }), {
        status: sent === big.length ? 201 : 202,
      });
    }) as unknown as typeof fetch;
    const out = await uploadLocalFile(graph, "/me/drive", { parentPath: "/", filename: "big.bin" }, uploads, "big.bin", {
      fetch: fetchFn,
      chunkSize: 3 * 320 * 1024,
    });
    expect(out).toEqual({ id: "big" });
    expect(received.length).toBe(5);
    expect(Buffer.concat(received).equals(big)).toBe(true);
  });

  it("stops, and cancels the session, if the file shrinks mid-upload", async () => {
    const graph = { api: () => ({ post: async () => ({ uploadUrl: "https://x.sharepoint.com/up" }) }) } as unknown as GraphClient;
    const methods: string[] = [];
    const fetchFn = (async (_url: string, init: RequestInit) => {
      methods.push(init.method!);
      return new Response(JSON.stringify({ nextExpectedRanges: ["327680-"] }), { status: 202 });
    }) as unknown as typeof fetch;
    const size = SIMPLE_UPLOAD_LIMIT + 1;
    const shrinking = { size, read: async (offset: number, length: number) => Buffer.alloc(offset === 0 ? length : 10) };
    await expect(uploadContent(graph, "/me/drive", { itemId: "i" }, shrinking, { fetch: fetchFn, chunkSize: 320 * 1024 })).rejects.toThrow(
      /changed during the upload/,
    );
    expect(methods).toEqual(["PUT", "DELETE"]);
  });
});

describe("files_upload", () => {
  const tool = findTool(filesTools, "files_upload");

  it("takes exactly one source, and a filename unless the file names itself", async () => {
    const ctx = makeContext({ uploadDir: uploads });
    await expect(callTool(tool, ctx, { parentPath: "/", filename: "a" })).rejects.toThrow(/exactly one/);
    await expect(callTool(tool, ctx, { parentPath: "/", filename: "a", contentBase64: "", localPath: "report.txt" })).rejects.toThrow(/exactly one/);
    await expect(callTool(tool, ctx, { parentPath: "/", contentBase64: "aGk=" })).rejects.toThrow(/filename is required/);
  });

  it("names the upload after the local file by default", async () => {
    const ctx = makeContext({ uploadDir: uploads });
    await callTool(tool, ctx, { parentPath: "/Docs", localPath: "sub/deep.txt" });
    expect(ctx.mock.calls[0]).toMatchObject({ method: "PUT", path: "/me/drive/root:/Docs/deep.txt:/content" });
  });

  it("says how to turn local uploads on when no folder is configured", async () => {
    const ctx = makeContext();
    await expect(callTool(tool, ctx, { parentPath: "/", localPath: "report.txt" })).rejects.toThrow(/set MCP_UPLOAD_DIR/);
    expect(ctx.mock.calls).toHaveLength(0);
  });
});
