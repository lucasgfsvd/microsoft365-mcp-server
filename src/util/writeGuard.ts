import type { ServerConfig, Surface, ToolDefinition } from "../types.js";
import { adminConsentGiven, availableInMode } from "../tools/appOnly.js";

/** Whether the tool is offered under the current config: not disabled, usable in this auth mode, writes allowed. */
export function isToolAllowed(tool: ToolDefinition, config: ServerConfig): boolean {
  if (config.disabledTools.has(tool.name)) return false;
  if (!availableInMode(tool, config)) return false;
  if (!adminConsentGiven(tool, config)) return false;
  if (!tool.mutating) return true;
  if (config.enableWrites) return true;
  const perSurface = config.perSurfaceWrites[tool.surface];
  return perSurface === true;
}

export function assertAllowed(tool: ToolDefinition, config: ServerConfig): void {
  if (isToolAllowed(tool, config)) return;
  if (config.disabledTools.has(tool.name)) throw new Error(`Tool "${tool.name}" is disabled by MCP_DISABLED_TOOLS.`);
  if (!availableInMode(tool, config)) {
    throw new Error(`Tool "${tool.name}" cannot be used with app-only access (client-credentials): Graph allows it only for a signed-in user.`);
  }
  if (!adminConsentGiven(tool, config)) {
    throw new Error(`Tool "${tool.name}" needs a permission only an admin can grant. Once your admin has approved it, set MCP_ENABLE_ADMIN_SCOPES=true and sign in again.`);
  }
  throw new WriteBlockedError(tool.name, tool.surface);
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
