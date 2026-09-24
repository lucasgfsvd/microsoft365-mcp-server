import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";

/** Above this, content goes through an upload session rather than one PUT. */
export const SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024;

/** Graph requires chunks in multiples of 320 KiB; ~10 MiB is its recommended size. */
const CHUNK_UNIT = 320 * 1024;
export const DEFAULT_CHUNK_SIZE = 32 * CHUNK_UNIT;

const MAX_ATTEMPTS = 4;

/** Where the bytes go: an existing item by id, or a path (created or replaced). */
export type UploadTarget = { itemId: string } | { parentPath: string; filename: string };

export interface UploadOptions {
  chunkSize?: number;
  /** Injectable for tests; defaults to the global fetch. */
  fetch?: typeof fetch;
  /** Injectable for tests; defaults to a real delay. */
  sleep?: (ms: number) => Promise<void>;
}

/** Graph address of the target, without the trailing `/content` or action. */
function address(base: string, target: UploadTarget): string {
  if ("itemId" in target) return `${base}/items/${target.itemId}`;
  return `${base}/root:${target.parentPath.replace(/\/+$/, "")}/${target.filename}:`;
}

/**
 * Write file content to OneDrive/SharePoint, whatever its size.
 *
 * Small files take the single-request PUT every tool used before. Larger ones
 * go through an upload session, which is the only way past Graph's simple-upload
 * limit. Returns the resulting driveItem either way.
 */
export async function uploadContent(
  graph: GraphClient,
  base: string,
  target: UploadTarget,
  content: Buffer,
  opts: UploadOptions = {},
): Promise<unknown> {
  if (content.byteLength <= SIMPLE_UPLOAD_LIMIT) {
    return graph.api(`${address(base, target)}/content`).put(content);
  }
  // Replace, matching what the simple PUT does to an existing file of that name.
  const body = "itemId" in target ? {} : { item: { "@microsoft.graph.conflictBehavior": "replace" } };
  const session = (await graph.api(`${address(base, target)}/createUploadSession`).post(body)) as {
    uploadUrl: string;
  };
  return sendChunks(session.uploadUrl, content, opts);
}

interface SessionStatus {
  nextExpectedRanges?: string[];
}

/** First byte the session still wants, per its `nextExpectedRanges` ("start-" or "start-end"). */
function nextOffset(status: SessionStatus): number | undefined {
  const first = status.nextExpectedRanges?.[0];
  return first === undefined ? undefined : Number(first.split("-")[0]);
}

function retryable(status: number): boolean {
  return status === 429 || status >= 500;
}

/**
 * PUT the content to a session's upload URL in chunks.
 *
 * The URL is pre-authorised and lives on a SharePoint host, so it is called with
 * plain fetch: the Graph client would attach the bearer token, which Microsoft
 * says must not be sent there. A failed chunk is retried after asking the session
 * where it actually stands, since a chunk can land even when its response is lost.
 */
async function sendChunks(uploadUrl: string, content: Buffer, opts: UploadOptions): Promise<unknown> {
  const doFetch = opts.fetch ?? fetch;
  const sleep = opts.sleep ?? ((ms: number) => new Promise((r) => setTimeout(r, ms)));
  const chunkSize = opts.chunkSize ?? DEFAULT_CHUNK_SIZE;
  if (chunkSize % CHUNK_UNIT !== 0) throw new Error(`chunkSize must be a multiple of ${CHUNK_UNIT} bytes.`);

  const total = content.byteLength;
  let offset = 0;
  let attempt = 0;
  try {
    for (;;) {
      const end = Math.min(offset + chunkSize, total) - 1;
      let res: Response | undefined;
      try {
        res = await doFetch(uploadUrl, {
          method: "PUT",
          headers: { "Content-Range": `bytes ${offset}-${end}/${total}` },
          body: content.subarray(offset, end + 1),
        });
      } catch {
        res = undefined; // network failure: treated like a retryable status
      }

      if (res && (res.status === 200 || res.status === 201)) return await res.json();
      if (res && res.status === 202) {
        const status = (await res.json()) as SessionStatus;
        offset = nextOffset(status) ?? end + 1;
        attempt = 0;
        continue;
      }
      if (res && !retryable(res.status) && res.status !== 416) {
        throw new Error(`Upload chunk ${offset}-${end} failed: HTTP ${res.status} ${await res.text()}`);
      }

      if (++attempt >= MAX_ATTEMPTS) {
        throw new Error(`Upload chunk ${offset}-${end} failed after ${MAX_ATTEMPTS} attempts.`);
      }
      const retryAfter = Number(res?.headers.get("Retry-After"));
      await sleep(Number.isFinite(retryAfter) && retryAfter > 0 ? retryAfter * 1000 : 2 ** attempt * 500);
      // Re-sync with the session before resending (416 means our range was already taken).
      const probe = await doFetch(uploadUrl, { method: "GET" }).catch(() => undefined);
      if (probe?.ok) offset = nextOffset((await probe.json()) as SessionStatus) ?? offset;
    }
  } catch (err) {
    // Best-effort cancel, so a failed upload does not linger as a session.
    await doFetch(uploadUrl, { method: "DELETE" }).catch(() => undefined);
    throw err;
  }
}
