import { tidyText } from "./text.js";

/**
 * Keys Graph fills with pre-authenticated URLs. `@microsoft.graph.downloadUrl`
 * (and its older name `@content.downloadUrl`) carries a `tempauth` token: anyone
 * holding the link can fetch the file, unauthenticated, for about an hour.
 * Tool results land in the model's context, transcripts and logs, and no tool
 * needs these links (downloads go by item id), so they never leave the server.
 * `downloadUrlNoAuth` is kept: it still requires a signed-in session.
 */
const PRE_AUTHENTICATED = new Set(["@microsoft.graph.downloadUrl", "@content.downloadUrl"]);

/**
 * `bodyPreview` is always plain text, and it is where newsletters put their
 * invisible padding (228 of one message's 255 characters). Tidying it here covers
 * every tool that returns messages: list, search, delta and raw batch results.
 */
const TIDY = new Set(["bodyPreview"]);

/**
 * Serialize a tool result for the client: drop pre-authenticated links and tidy
 * message previews, at any depth.
 */
export function serializeResult(result: unknown): string {
  // A tool whose Graph call answered 204 No Content (a Planner PATCH, for one)
  // resolves to undefined, and JSON.stringify(undefined) is not a string: the
  // result became an invalid MCP message and the client reported a failure for
  // an action that had succeeded. Seen live.
  if (result === undefined) return JSON.stringify({ ok: true }, null, 2);
  return JSON.stringify(
    result,
    (key, value) => {
      if (PRE_AUTHENTICATED.has(key)) return undefined;
      if (TIDY.has(key) && typeof value === "string") return tidyText(value);
      return value;
    },
    2,
  );
}
