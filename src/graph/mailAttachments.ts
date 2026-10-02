import { promises as fs } from "node:fs";
import path from "node:path";
import { z } from "zod";
import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import type { ToolContext } from "../types.js";
import { Filename } from "../util/schema.js";
import { downloadInline, getItemMeta } from "./download.js";
import { resolveUploadPath, fileSource } from "./localUpload.js";
import { bufferSource, sendChunks, type ContentSource } from "./upload.js";

/** One file to attach, from one of three places. */
export const AttachmentInput = z
  .object({
    localPath: z.string().min(1).optional().describe("A file in MCP_UPLOAD_DIR, relative to it."),
    driveItemId: z.string().optional().describe("A OneDrive or SharePoint file, by item id."),
    driveId: z.string().optional().describe("The drive holding driveItemId; omit for your OneDrive."),
    contentBase64: z.string().optional().describe("The bytes, base64-encoded; needs name."),
    name: Filename.optional().describe("File name in the message. Defaults to the source file's name."),
  })
  .refine((a) => [a.localPath, a.driveItemId, a.contentBase64].filter((x) => x !== undefined).length === 1, {
    message: "Each attachment needs exactly one of localPath, driveItemId or contentBase64",
  })
  .refine((a) => a.contentBase64 === undefined || a.name !== undefined, { message: "An attachment from contentBase64 needs a name" });
export type AttachmentInput = z.infer<typeof AttachmentInput>;

/** Below this an attachment goes in one request; above, through an upload session (to 150 MB). */
export const SIMPLE_ATTACHMENT_LIMIT = 3 * 1024 * 1024;
/** Outlook takes ranges of up to 4 MB; a multiple of 320 KiB keeps the chunk rules of drive sessions. */
const ATTACHMENT_CHUNK = 12 * 320 * 1024;

interface Resolved {
  name: string;
  source: ContentSource;
  close: () => Promise<void>;
}

async function resolve(graph: GraphClient, a: AttachmentInput, uploadDir: string | undefined): Promise<Resolved> {
  if (a.localPath !== undefined) {
    if (!uploadDir) throw new Error("Attaching from disk is off: set MCP_UPLOAD_DIR to the one folder the server may read from.");
    const real = await resolveUploadPath(uploadDir, a.localPath);
    const handle = await fs.open(real, "r");
    const { size } = await handle.stat();
    return { name: a.name ?? path.basename(a.localPath), source: fileSource(handle, size), close: () => handle.close() };
  }
  if (a.driveItemId !== undefined) {
    const item = `${a.driveId ? `/drives/${encodeURIComponent(a.driveId)}` : "/me/drive"}/items/${encodeURIComponent(a.driveItemId)}`;
    const meta = await getItemMeta(graph, item);
    if (meta.folder) throw new Error(`${meta.name} is a folder; attach files.`);
    const bytes = await downloadInline(graph, `${item}/content`);
    return { name: a.name ?? meta.name, source: bufferSource(bytes), close: async () => undefined };
  }
  return { name: a.name!, source: bufferSource(Buffer.from(a.contentBase64 ?? "", "base64")), close: async () => undefined };
}

/**
 * Attach files to a draft at `messagePath` (e.g. /me/messages/{id}). Small ones
 * are posted whole; larger ones go through an attachment upload session, read a
 * chunk at a time, so a file from disk is never held in memory.
 */
export async function attachFiles(
  graph: GraphClient,
  messagePath: string,
  attachments: AttachmentInput[],
  ctx: Pick<ToolContext, "config" | "signal">,
): Promise<Array<{ name: string; size: number }>> {
  const done: Array<{ name: string; size: number }> = [];
  for (const a of attachments) {
    ctx.signal?.throwIfAborted();
    const r = await resolve(graph, a, ctx.config.uploadDir);
    try {
      const { size } = r.source;
      if (size < SIMPLE_ATTACHMENT_LIMIT) {
        await graph.api(`${messagePath}/attachments`).post({
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: r.name,
          contentBytes: (await r.source.read(0, size)).toString("base64"),
        });
      } else {
        const session = (await graph
          .api(`${messagePath}/attachments/createUploadSession`)
          .post({ AttachmentItem: { attachmentType: "file", name: r.name, size } })) as { uploadUrl: string };
        await sendChunks(session.uploadUrl, r.source, { chunkSize: ATTACHMENT_CHUNK, signal: ctx.signal });
      }
      done.push({ name: r.name, size });
    } finally {
      await r.close();
    }
  }
  return done;
}
