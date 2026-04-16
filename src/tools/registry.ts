import type { ToolDefinition, Surface, ServerConfig } from "../types.js";
import { isToolAllowed } from "../util/writeGuard.js";

export class ToolRegistry {
  private readonly tools = new Map<string, ToolDefinition>();

  register(tool: ToolDefinition): void {
    if (this.tools.has(tool.name)) {
      throw new Error(`Duplicate tool name: ${tool.name}`);
    }
    this.tools.set(tool.name, tool);
  }

  registerAll(tools: ToolDefinition[]): void {
    for (const t of tools) this.register(t);
  }

  list(config: ServerConfig): ToolDefinition[] {
    return Array.from(this.tools.values()).filter((t) => isToolAllowed(t, config));
  }

  get(name: string): ToolDefinition | undefined {
    return this.tools.get(name);
  }

  bySurface(surface: Surface): ToolDefinition[] {
    return Array.from(this.tools.values()).filter((t) => t.surface === surface);
  }
}
