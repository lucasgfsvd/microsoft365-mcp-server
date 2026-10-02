import type { ToolDefinition } from "../../types.js";
import { todoTools } from "./todo.js";
import { plannerTools } from "./planner.js";
import { plannerDetailTools } from "./plannerDetails.js";

export const tasksTools: ToolDefinition[] = [...todoTools, ...plannerTools, ...plannerDetailTools];
