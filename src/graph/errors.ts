import { GraphError } from "@microsoft/microsoft-graph-client";

export class M365McpError extends Error {
  constructor(
    message: string,
    public readonly code: string,
    public readonly statusCode?: number,
    public readonly requestId?: string,
  ) {
    super(message);
    this.name = "M365McpError";
  }
}

export function normalizeGraphError(err: unknown): M365McpError {
  if (err instanceof GraphError) {
    return new M365McpError(
      err.message || "Graph request failed",
      err.code || "GraphError",
      err.statusCode,
      err.requestId ?? undefined,
    );
  }
  if (err instanceof Error) {
    return new M365McpError(err.message, err.name || "Error");
  }
  return new M365McpError(String(err), "UnknownError");
}
