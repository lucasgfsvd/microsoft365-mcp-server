import { z } from "zod";

const Day = z.enum(["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"]);
const DAYS = Day.options;

/**
 * A repeating event, as a person would say it: "every weekday", "every other
 * Tuesday until June", "the first Monday of each month, 10 times". Anything left
 * out is taken from the event's start date.
 */
export const RecurrenceInput = z
  .object({
    pattern: z.enum(["daily", "weekly", "absoluteMonthly", "relativeMonthly", "absoluteYearly"]),
    interval: z.number().int().min(1).default(1).describe("Every N days/weeks/months/years."),
    daysOfWeek: z.array(Day).optional().describe("weekly and relativeMonthly; default: the start's weekday."),
    dayOfMonth: z.number().int().min(1).max(31).optional().describe("absoluteMonthly and absoluteYearly; default: the start's day."),
    weekIndex: z.enum(["first", "second", "third", "fourth", "last"]).optional().describe("relativeMonthly, e.g. the 'first' Monday."),
    until: z.iso.date().optional().describe("Last date on which it may occur."),
    occurrences: z.number().int().min(1).max(999).optional().describe("Number of times, instead of until."),
  })
  .refine((r) => !(r.until && r.occurrences), { message: "Give until or occurrences, not both" })
  .refine((r) => r.pattern !== "relativeMonthly" || r.weekIndex, { message: "relativeMonthly needs weekIndex" });
export type RecurrenceInput = z.infer<typeof RecurrenceInput>;

/** Graph's patternedRecurrence for an event starting at `start` (local date-time) in `timeZone`. */
export function toPatternedRecurrence(r: RecurrenceInput, start: string, timeZone: string) {
  const startDate = start.slice(0, 10);
  const d = new Date(`${startDate}T00:00:00Z`);
  if (Number.isNaN(d.getTime())) throw new Error(`Cannot read the start date of "${start}".`);
  const weekday = DAYS[d.getUTCDay()]!;
  const pattern: Record<string, unknown> = { type: r.pattern, interval: r.interval };
  if (r.pattern === "weekly" || r.pattern === "relativeMonthly") pattern.daysOfWeek = r.daysOfWeek ?? [weekday];
  if (r.pattern === "relativeMonthly") pattern.index = r.weekIndex;
  if (r.pattern === "absoluteMonthly" || r.pattern === "absoluteYearly") pattern.dayOfMonth = r.dayOfMonth ?? d.getUTCDate();
  if (r.pattern === "absoluteYearly") pattern.month = d.getUTCMonth() + 1;

  const range = r.until
    ? { type: "endDate", startDate, endDate: r.until }
    : r.occurrences
      ? { type: "numbered", startDate, numberOfOccurrences: r.occurrences }
      : { type: "noEnd", startDate };
  return { pattern, range: { ...range, recurrenceTimeZone: timeZone } };
}
