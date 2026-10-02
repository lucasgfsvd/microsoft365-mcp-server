import type { ToolDefinition } from "../../types.js";
import { todoTools } from "./todo.js";
import { plannerTools } from "./planner.js";

export const tasksTools: ToolDefinition[] = [...todoTools, ...plannerTools];
