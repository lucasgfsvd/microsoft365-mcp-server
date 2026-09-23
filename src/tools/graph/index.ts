import type { ToolDefinition } from "../../types.js";
import { graphBatchTools } from "./batch.js";
import { graphSearchTools } from "./search.js";

/** Cross-cutting Graph capabilities that are not tied to one surface. */
export const graphTools: ToolDefinition[] = [...graphBatchTools, ...graphSearchTools];
