import type { ServerConfig, Surface, ToolDefinition } from "../types.js";

/** Returns true if this tool's mutation should be allowed under current config. */
export function isToolAllowed(tool: ToolDefinition, config: ServerConfig): boolean {
  if (config.disabledTools.has(tool.name)) return false;
  if (!tool.mutating) return true;
  if (config.enableWrites) return true;
  const perSurface = config.perSurfaceWrites[tool.surface];
  return perSurface === true;
}

export function assertAllowed(tool: ToolDefinition, config: ServerConfig): void {
  if (!isToolAllowed(tool, config)) {
    throw new WriteBlockedError(tool.name, tool.surface);
  }
}

export class WriteBlockedError extends Error {
  constructor(public readonly toolName: string, public readonly surface: Surface) {
    super(
      `Tool "${toolName}" is a write operation and is disabled. Enable with --enable-writes, ` +
        `MCP_ENABLE_WRITES=true, or MCP_ENABLE_${surface.toUpperCase()}_WRITE=true.`,
    );
    this.name = "WriteBlockedError";
  }
}
