import { createHash } from "node:crypto";
import { promises as fs } from "node:fs";
import path from "node:path";
import { Readable, Transform } from "node:stream";
import { pipeline } from "node:stream/promises";
import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";

/**
 * Largest file returned inside a tool result. Base64 inflates it by a third,
 * and the result has to fit the *client's* inbound MCP message limit — 10 MiB
 * by the SDK's default — or the client drops the connection.
 */
export const INLINE_LIMIT = 5 * 1024 * 1024;

export interface DriveItemMeta {
  name: string;
  size: number;
  file?: { mimeType?: string };
  folder?: unknown;
}

export async function getItemMeta(graph: GraphClient, itemUrl: string): Promise<DriveItemMeta> {
  return (await graph.api(itemUrl).select("name,size,file,folder").get()) as DriveItemMeta;
}

/** Whatever getStream() hands back, as a Node Readable (it is a web stream under native fetch). */
function toReadable(stream: unknown): Readable {
  if (stream instanceof Readable) return stream;
  if (stream && typeof (stream as ReadableStream).getReader === "function") {
    return Readable.fromWeb(stream as import("node:stream/web").ReadableStream);
  }
  return Readable.from(stream as AsyncIterable<Buffer>);
}

const TEXT_TYPES = /^(text\/|application\/(json|xml|javascript|x-yaml|yaml|csv|x-sh|sql)\b|.*\+(json|xml)$)/;

export type InlineContent =
  | { encoding: "utf8"; text: string; byteLength: number; mimeType?: string }
  | { encoding: "base64"; base64: string; byteLength: number; mimeType?: string };

/**
 * Bytes for the tool result: plain text when the type says text and the bytes
 * really are UTF-8, since base64 of text costs a third more context for nothing.
 */
export function toInline(bytes: Buffer, mimeType?: string): InlineContent {
  if (mimeType && TEXT_TYPES.test(mimeType)) {
    try {
      const text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
      return { encoding: "utf8", text, byteLength: bytes.byteLength, mimeType };
    } catch {
      // Labelled text but not valid UTF-8: fall through to base64.
    }
  }
  return { encoding: "base64", base64: bytes.toString("base64"), byteLength: bytes.byteLength, mimeType };
}

export async function downloadInline(graph: GraphClient, contentUrl: string): Promise<Buffer> {
  const chunks: Buffer[] = [];
  for await (const c of toReadable(await graph.api(contentUrl).getStream())) {
    chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c));
  }
  return Buffer.concat(chunks);
}

const WINDOWS_RESERVED = /^(con|prn|aux|nul|com[0-9]|lpt[0-9])(\..*)?$/i;

/**
 * A drive item's name, made safe to create inside the download folder: no path
 * separators or traversal, no characters Windows refuses, no reserved device
 * names. The name comes from the drive, which other people can write to.
 */
export function safeFileName(name: string): string {
  let n = name
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_")
    .replace(/[. ]+$/, "")
    .trim();
  if (!n || n === "." || n === "..") n = "download";
  if (WINDOWS_RESERVED.test(n)) n = `_${n}`;
  return n.slice(0, 200);
}

/** "report.pdf" → "report (1).pdf", "report (2).pdf", … */
function candidate(name: string, i: number): string {
  if (i === 0) return name;
  const ext = path.extname(name);
  return `${name.slice(0, name.length - ext.length)} (${i})${ext}`;
}

export interface SavedFile {
  path: string;
  byteLength: number;
  sha256: string;
}

/**
 * Stream a file to disk without holding it in memory.
 *
 * Opens the target with the exclusive-create flag, so an existing file — or a
 * symlink planted under the same name — is never written through; a taken name
 * moves on to "name (1)". A failed transfer removes its partial file.
 */
export async function downloadToDir(
  graph: GraphClient,
  contentUrl: string,
  dir: string,
  itemName: string,
): Promise<SavedFile> {
  await fs.mkdir(dir, { recursive: true });
  const root = path.resolve(dir);
  const base = safeFileName(itemName);

  let handle: fs.FileHandle | undefined;
  let target = "";
  for (let i = 0; i < 1000 && !handle; i++) {
    target = path.join(root, candidate(base, i));
    if (path.dirname(target) !== root) throw new Error(`Refusing to write outside ${root}.`);
    try {
      handle = await fs.open(target, "wx");
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code !== "EEXIST") throw err;
    }
  }
  if (!handle) throw new Error(`No free name for ${base} in ${root}.`);

  const hash = createHash("sha256");
  let byteLength = 0;
  const tap = new Transform({
    transform(chunk: Buffer, _enc, done) {
      hash.update(chunk);
      byteLength += chunk.length;
      done(null, chunk);
    },
  });

  try {
    const source = toReadable(await graph.api(contentUrl).getStream());
    await pipeline(source, tap, handle.createWriteStream());
  } catch (err) {
    await handle.close().catch(() => undefined);
    await fs.unlink(target).catch(() => undefined);
    throw err;
  }
  return { path: target, byteLength, sha256: hash.digest("hex") };
}
