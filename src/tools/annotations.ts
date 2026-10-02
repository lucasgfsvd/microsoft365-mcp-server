import type { ToolDefinition } from "../types.js";

/**
 * MCP tool annotations: the standard hints a client uses to decide what to ask
 * the user before running a tool. Clients may treat them as untrusted; they
 * guide confirmation prompts, they do not enforce anything (writeGuard does).
 */
export interface ToolAnnotations {
  readOnlyHint: boolean;
  /** A write that can change or remove what already exists (false: only adds). */
  destructiveHint?: boolean;
  /** Calling again with the same arguments has no further effect. */
  idempotentHint?: boolean;
  /** Reaches other people: mail sent, invitations, posts, shared plans, links. */
  openWorldHint: boolean;
}

type WriteKind = Required<Pick<ToolAnnotations, "destructiveHint" | "idempotentHint">>;

/** How a write behaves, by the verb in its name (surface_verb_object). */
const BY_VERB: Array<[RegExp, WriteKind]> = [
  [/^(delete|update|set|clear|rename|replace|complete)$/, { destructiveHint: true, idempotentHint: true }],
  // Writes a file at a path, replacing one already there.
  [/^(upload|copy)$/, { destructiveHint: true, idempotentHint: true }],
  [/^(create|add|append|insert|post|reply|send)$/, { destructiveHint: false, idempotentHint: false }],
  // A sharing link of a given type is returned again, not duplicated.
  [/^share$/, { destructiveHint: false, idempotentHint: true }],
  [/^run$/, { destructiveHint: false, idempotentHint: true }],
];

/** Where the verb alone gets it wrong. */
const EXCEPTIONS: Record<string, WriteKind> = {
  // "Create" a file at a path: an existing file of that name is replaced.
  excel_create_workbook: { destructiveHint: true, idempotentHint: true },
  excel_create_from_template: { destructiveHint: true, idempotentHint: true },
  word_create_document: { destructiveHint: true, idempotentHint: true },
  word_create_from_template: { destructiveHint: true, idempotentHint: true },
  powerpoint_create_deck: { destructiveHint: true, idempotentHint: true },
  powerpoint_create_from_template: { destructiveHint: true, idempotentHint: true },
  // By position: a second call removes the next paragraph or slide.
  word_delete_paragraph: { destructiveHint: true, idempotentHint: false },
  powerpoint_delete_slide: { destructiveHint: true, idempotentHint: false },
  // Renames on a name clash, so a second call makes a second folder.
  files_create_folder: { destructiveHint: false, idempotentHint: false },
};

/** Writes other people see or receive. */
const OPEN_WORLD = /^(mail_(send|reply)_|teams_(post|reply)_|calendar_(create|update|delete)_|planner_(create|complete|delete)_|files_share$)/;

export function writeKind(name: string): WriteKind | undefined {
  if (EXCEPTIONS[name]) return EXCEPTIONS[name];
  const verb = name.split("_")[1] ?? "";
  return BY_VERB.find(([re]) => re.test(verb))?.[1];
}

export function annotationsFor(tool: Pick<ToolDefinition, "name" | "mutating">): ToolAnnotations {
  if (!tool.mutating) return { readOnlyHint: true, openWorldHint: false };
  const kind = writeKind(tool.name);
  // An unclassified write is assumed to be the risky kind.
  return { readOnlyHint: false, ...(kind ?? { destructiveHint: true, idempotentHint: false }), openWorldHint: OPEN_WORLD.test(tool.name) };
}
