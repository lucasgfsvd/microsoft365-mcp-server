import type { ServerConfig, ToolDefinition } from "../types.js";

/**
 * Tools Graph does not allow with an app-only token, hidden in
 * client-credentials mode rather than offered and left to fail:
 *
 * - OneNote: Microsoft ended app-only access to the OneNote API in 2025.
 * - Posting to Teams: application permissions only cover importing history,
 *   not sending as a person.
 * - Microsoft Search: application permissions do not cover mail or files.
 */
export const DELEGATED_ONLY = new Set([
  "onenote_list_notebooks",
  "onenote_list_sections",
  "onenote_list_pages",
  "onenote_get_page_content",
  "onenote_create_notebook",
  "onenote_create_section",
  "onenote_create_page",
  "onenote_delete_page",
  "teams_post_channel_message",
  "teams_reply_channel_message",
  "teams_post_chat_message",
  "teams_send_direct_message",
  "teams_update_chat_message",
  "teams_delete_chat_message",
  "graph_search",
]);

export function availableInMode(tool: Pick<ToolDefinition, "name">, config: Pick<ServerConfig, "authMode">): boolean {
  return config.authMode !== "client-credentials" || !DELEGATED_ONLY.has(tool.name);
}

/** A tool needing an admin-granted scope is offered once the user opts in to those scopes. */
export function adminConsentGiven(tool: Pick<ToolDefinition, "needsAdminConsent">, config: Pick<ServerConfig, "authMode" | "adminScopes">): boolean {
  return !tool.needsAdminConsent || config.authMode === "client-credentials" || config.adminScopes === true;
}
