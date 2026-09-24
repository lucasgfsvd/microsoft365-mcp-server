/**
 * MCP prompt templates: user-invoked workflows that expand into instructions
 * for the model. They carry what a caller would otherwise have to rediscover —
 * exact tool names, pre-computed Graph queries, and the rules learned the hard
 * way (batch independent reads, never send mail, search types that cannot mix).
 */

export interface PromptArgument {
  name: string;
  description: string;
  required?: boolean;
}

export interface PromptContext {
  /** Tools visible under the current config; prompts adapt to what is enabled. */
  tools: ReadonlySet<string>;
  now: Date;
  /** IANA zone the server runs in, e.g. "Europe/Paris". */
  timeZone: string;
}

export interface PromptDefinition {
  name: string;
  title: string;
  description: string;
  arguments: PromptArgument[];
  /** Listed only when every one of these is visible. */
  requiredTools: string[];
  render(args: Record<string, string | undefined>, ctx: PromptContext): string;
}

/** Shared guardrail text: every workflow here reads freely and writes nothing without asking. */
export const NO_SENDING =
  "Do not send, reply to, delete or move anything. mail_reply_message and mail_send_message send " +
  "immediately, so never call them here; if a reply is worth writing, draft it with mail_create_draft " +
  "(when available) and say it is waiting in Drafts.";

/**
 * Parse a lookback like "24h", "3d" or an ISO date into an instant. Unparseable
 * input throws, so the user learns immediately rather than getting a wrong window.
 */
export function parseSince(since: string | undefined, now: Date, fallbackHours = 24): Date {
  if (!since) return new Date(now.getTime() - fallbackHours * 3_600_000);
  const rel = /^\s*(\d+)\s*([hd])\s*$/i.exec(since);
  if (rel) return new Date(now.getTime() - Number(rel[1]) * (rel[2]?.toLowerCase() === "h" ? 3_600_000 : 86_400_000));
  const abs = new Date(since);
  if (Number.isNaN(abs.getTime())) throw new Error(`Cannot read "${since}" as a lookback: use e.g. "24h", "3d" or a date like 2026-09-01.`);
  return abs;
}

/** Today's date in `timeZone`, as YYYY-MM-DD. */
export function localDate(now: Date, timeZone: string): string {
  const p = new Intl.DateTimeFormat("en-CA", { timeZone, year: "numeric", month: "2-digit", day: "2-digit" })
    .formatToParts(now)
    .reduce<Record<string, string>>((acc, x) => ((acc[x.type] = x.value), acc), {});
  return `${p.year}-${p.month}-${p.day}`;
}

/** Start and end of the day `ymd` in `timeZone`, as UTC ISO strings. */
export function dayWindow(ymd: string, timeZone: string): { start: string; end: string } {
  return { start: zonedMidnight(ymd, timeZone).toISOString(), end: zonedMidnight(nextDate(ymd), timeZone).toISOString() };
}

function nextDate(ymd: string): string {
  const d = new Date(`${ymd}T00:00:00Z`);
  d.setUTCDate(d.getUTCDate() + 1);
  return d.toISOString().slice(0, 10);
}

/** The UTC instant of 00:00 on `ymd` in `timeZone` (handles DST by measuring the zone's offset then). */
function zonedMidnight(ymd: string, timeZone: string): Date {
  const guess = new Date(`${ymd}T00:00:00Z`);
  const shown = new Date(guess.toLocaleString("en-US", { timeZone }));
  const utc = new Date(guess.toLocaleString("en-US", { timeZone: "UTC" }));
  return new Date(guess.getTime() - (shown.getTime() - utc.getTime()));
}

/** Validate "YYYY-MM-DD", defaulting to today in `timeZone`. */
export function parseDay(date: string | undefined, now: Date, timeZone: string): string {
  if (!date) return localDate(now, timeZone);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || Number.isNaN(new Date(`${date}T00:00:00Z`).getTime())) {
    throw new Error(`Expected a date like 2026-09-24, got "${date}".`);
  }
  return date;
}
