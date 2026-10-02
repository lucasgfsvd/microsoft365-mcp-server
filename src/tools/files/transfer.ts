import type { ToolContext } from "../../types.js";
import type { UploadOptions } from "../../graph/upload.js";
import type { DownloadOptions } from "../../graph/download.js";
import { MB } from "../../util/progress.js";

/** A tool call's cancellation and progress, as an upload takes them. */
export function uploadOptions(ctx: Pick<ToolContext, "signal" | "progress">, name: string): UploadOptions {
  const progress = ctx.progress;
  return {
    signal: ctx.signal,
    onProgress: progress && ((sent, total) => progress(sent, total, `Uploading ${name}: ${MB(sent)} of ${MB(total)}`)),
  };
}

/** The same for a download of `size` bytes to disk. */
export function downloadOptions(ctx: Pick<ToolContext, "signal" | "progress">, name: string, size: number): DownloadOptions {
  const progress = ctx.progress;
  return {
    signal: ctx.signal,
    onProgress: progress && ((written) => progress(written, size, `Downloading ${name}: ${MB(written)} of ${MB(size)}`)),
  };
}
