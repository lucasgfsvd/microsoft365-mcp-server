import { promises as fs } from "node:fs";
import path from "node:path";
import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { uploadContent, type ContentSource, type UploadOptions, type UploadTarget } from "./upload.js";

/**
 * Resolve `localPath` to a regular file inside the upload folder, or refuse.
 *
 * A path comes from the model, and so possibly from whatever text it read, so
 * it is confined to the one folder the user set aside: relative paths resolve
 * inside it, absolute ones must already point inside it, and the check runs on
 * the real path, so a symlink in the folder cannot lead out of it.
 */
export async function resolveUploadPath(uploadDir: string, localPath: string): Promise<string> {
  if (localPath.includes("\0")) throw new Error("localPath cannot contain NUL characters.");
  const root = await fs.realpath(uploadDir).catch(() => {
    throw new Error(`The upload folder ${uploadDir} does not exist.`);
  });
  const outside = new Error(`${localPath} is outside the upload folder ${root}; only files inside it can be uploaded.`);
  // Checked before touching the disk, so a path outside cannot even be probed for existence.
  // The folder as configured or as resolved, should it be a link itself.
  const lexical = path.resolve(root, localPath);
  if (!inside(root, lexical) && !inside(path.resolve(uploadDir), lexical)) throw outside;
  const real = await fs.realpath(lexical).catch(() => {
    throw new Error(`No file at ${localPath} in the upload folder ${root}.`);
  });
  if (!inside(root, real)) throw outside;
  if (!(await fs.stat(real)).isFile()) throw new Error(`${localPath} is not a file.`);
  return real;
}

function inside(root: string, p: string): boolean {
  const rel = path.relative(root, p);
  return rel !== "" && rel !== ".." && !rel.startsWith(`..${path.sep}`) && !path.isAbsolute(rel);
}

/** A source reading ranges from an open file. */
export function fileSource(handle: fs.FileHandle, size: number): ContentSource {
  return {
    size,
    async read(offset, length) {
      const buf = Buffer.alloc(Math.min(length, size - offset));
      let filled = 0;
      while (filled < buf.byteLength) {
        const { bytesRead } = await handle.read(buf, filled, buf.byteLength - filled, offset + filled);
        if (bytesRead === 0) break; // shorter than when we started: the caller reports it
        filled += bytesRead;
      }
      return buf.subarray(0, filled);
    },
  };
}

/**
 * Upload a file from the upload folder, streamed from disk a chunk at a time:
 * neither the bytes nor their base64 pass through the conversation, and size is
 * limited only by OneDrive.
 */
export async function uploadLocalFile(
  graph: GraphClient,
  base: string,
  target: UploadTarget,
  uploadDir: string,
  localPath: string,
  opts: UploadOptions = {},
): Promise<unknown> {
  const real = await resolveUploadPath(uploadDir, localPath);
  const handle = await fs.open(real, "r");
  try {
    const { size } = await handle.stat();
    return await uploadContent(graph, base, target, fileSource(handle, size), opts);
  } finally {
    await handle.close();
  }
}
