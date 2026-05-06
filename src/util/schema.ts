import { z } from "zod";

export const PaginationInput = z.object({
  top: z.number().int().min(1).max(999).optional().describe("Page size (1–999)."),
  skip: z.number().int().min(0).optional().describe("Items to skip."),
  nextLink: z.string().url().optional().describe("Opaque @odata.nextLink cursor from a prior call."),
});

export type PaginationInput = z.infer<typeof PaginationInput>;

export const IdInput = z.object({ id: z.string().min(1) });

function hasControlChar(s: string): boolean {
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 0x20 || c === 0x7f) return true;
  }
  return false;
}

/**
 * Drive-relative path, e.g. "/Reports/2026" or "/Templates/proposal.docx".
 * Must start with "/", cannot traverse upward (no ".." segments), and cannot
 * contain URL-control characters that would corrupt the Graph URL we splice
 * it into.
 */
export const DrivePath = z
  .string()
  .min(1)
  .refine((p) => p.startsWith("/"), { message: "Path must start with '/'" })
  .refine((p) => !p.split("/").includes(".."), { message: "Path cannot contain '..' segments" })
  .refine((p) => !hasControlChar(p), { message: "Path cannot contain control characters" })
  .refine((p) => !/[?#]/.test(p), { message: "Path cannot contain '?' or '#'" });

/** Filename without directory separators or path-traversal sequences. */
export const Filename = z
  .string()
  .min(1)
  .refine((f) => !f.includes("/") && !f.includes("\\"), { message: "Filename cannot contain '/' or '\\'" })
  .refine((f) => f !== "." && f !== "..", { message: "Filename cannot be '.' or '..'" })
  .refine((f) => !hasControlChar(f), { message: "Filename cannot contain control characters" });

export function trimEmpty<T extends Record<string, unknown>>(obj: T): Partial<T> {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(obj)) {
    if (v !== undefined && v !== null && v !== "") out[k] = v;
  }
  return out as Partial<T>;
}
