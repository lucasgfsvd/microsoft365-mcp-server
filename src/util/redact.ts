/**
 * Keys Graph fills with pre-authenticated URLs. `@microsoft.graph.downloadUrl`
 * (and its older name `@content.downloadUrl`) carries a `tempauth` token: anyone
 * holding the link can fetch the file, unauthenticated, for about an hour.
 * Tool results land in the model's context, transcripts and logs, and no tool
 * needs these links (downloads go by item id), so they never leave the server.
 * `downloadUrlNoAuth` is kept: it still requires a signed-in session.
 */
const PRE_AUTHENTICATED = new Set(["@microsoft.graph.downloadUrl", "@content.downloadUrl"]);

/** Serialize a tool result for the client, dropping pre-authenticated links at any depth. */
export function serializeResult(result: unknown): string {
  return JSON.stringify(result, (key, value) => (PRE_AUTHENTICATED.has(key) ? undefined : value), 2);
}
