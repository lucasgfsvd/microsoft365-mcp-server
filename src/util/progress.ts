/** Reports how far a long operation has got: `done` of `total` units (bytes, for transfers). */
export type ProgressReporter = (done: number, total?: number, message?: string) => void;

export interface ProgressNotification {
  method: "notifications/progress";
  params: { progressToken: string | number; progress: number; total?: number; message?: string };
}

/**
 * A reporter that sends MCP progress notifications for one request, or none
 * when the client did not ask (no progressToken). Updates are throttled, since
 * a transfer can report every few kilobytes; the first and the last always go.
 */
export function progressReporter(
  token: string | number | undefined,
  send: (n: ProgressNotification) => Promise<void>,
  { intervalMs = 250, now = () => Date.now() }: { intervalMs?: number; now?: () => number } = {},
): ProgressReporter | undefined {
  if (token === undefined) return undefined;
  let last = -Infinity;
  let lastDone = -1;
  return (done, total, message) => {
    const final = total !== undefined && done >= total;
    // Progress must increase; and between the first and the last, one per interval.
    if (done <= lastDone || (!final && now() - last < intervalMs)) return;
    last = now();
    lastDone = done;
    send({ method: "notifications/progress", params: { progressToken: token, progress: done, total, message } }).catch(() => undefined);
  };
}

export const MB = (n: number) => `${(n / 1024 / 1024).toFixed(1)} MB`;
