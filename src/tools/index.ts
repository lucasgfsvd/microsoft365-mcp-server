import type { ToolDefinition } from "../types.js";
import { mailTools } from "./mail/index.js";
import { calendarTools } from "./calendar/index.js";
import { contactsTools } from "./contacts/index.js";
import { filesTools } from "./files/index.js";
import { teamsTools } from "./teams/index.js";
import { tasksTools } from "./tasks/index.js";
import { onenoteTools } from "./onenote/index.js";
import { excelTools } from "./excel/index.js";
import { wordTools } from "./word/index.js";
import { powerpointTools } from "./powerpoint/index.js";

export function allTools(): ToolDefinition[] {
  return [
    ...mailTools,
    ...calendarTools,
    ...contactsTools,
    ...filesTools,
    ...teamsTools,
    ...tasksTools,
    ...onenoteTools,
    ...excelTools,
    ...wordTools,
    ...powerpointTools,
  ];
}
