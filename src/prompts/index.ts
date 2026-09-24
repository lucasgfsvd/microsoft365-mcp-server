import type { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { GetPromptRequestSchema, ListPromptsRequestSchema } from "@modelcontextprotocol/sdk/types.js";
import { dailyBrief } from "./dailyBrief.js";
import { inboxTriage } from "./inboxTriage.js";
import { meetingPrep } from "./meetingPrep.js";
import type { PromptContext, PromptDefinition } from "./types.js";

export const PROMPTS: PromptDefinition[] = [dailyBrief, inboxTriage, meetingPrep];

/** Prompts whose tools are all enabled under the current config. */
export function availablePrompts(tools: ReadonlySet<string>): PromptDefinition[] {
  return PROMPTS.filter((p) => p.requiredTools.every((t) => tools.has(t)));
}

export function renderPrompt(name: string, args: Record<string, string | undefined>, ctx: PromptContext): string {
  const prompt = availablePrompts(ctx.tools).find((p) => p.name === name);
  if (!prompt) throw new Error(`Unknown or unavailable prompt: ${name}`);
  for (const a of prompt.arguments) {
    if (a.required && !args[a.name]) throw new Error(`Prompt ${name} needs argument "${a.name}".`);
  }
  return prompt.render(args, ctx);
}

/**
 * Serve the prompts over MCP. `visibleTools` is read per request so the list
 * always matches what tools/list shows.
 */
export function registerPrompts(server: Server, visibleTools: () => ReadonlySet<string>): void {
  const context = (): PromptContext => ({
    tools: visibleTools(),
    now: new Date(),
    timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
  });

  server.setRequestHandler(ListPromptsRequestSchema, async () => ({
    prompts: availablePrompts(visibleTools()).map((p) => ({
      name: p.name,
      title: p.title,
      description: p.description,
      arguments: p.arguments,
    })),
  }));

  server.setRequestHandler(GetPromptRequestSchema, async (req) => {
    const ctx = context();
    const text = renderPrompt(req.params.name, req.params.arguments ?? {}, ctx);
    const prompt = PROMPTS.find((p) => p.name === req.params.name);
    return {
      description: prompt?.description,
      messages: [{ role: "user" as const, content: { type: "text" as const, text } }],
    };
  });
}
