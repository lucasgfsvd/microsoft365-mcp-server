import { z } from "zod";

export const PaginationInput = z.object({
  top: z.number().int().min(1).max(999).optional().describe("Page size (1–999)."),
  skip: z.number().int().min(0).optional().describe("Items to skip."),
  nextLink: z.string().url().optional().describe("Opaque @odata.nextLink cursor from a prior call."),
});

export type PaginationInput = z.infer<typeof PaginationInput>;

export const IdInput = z.object({ id: z.string().min(1) });

export function trimEmpty<T extends Record<string, unknown>>(obj: T): Partial<T> {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(obj)) {
    if (v !== undefined && v !== null && v !== "") out[k] = v;
  }
  return out as Partial<T>;
}
