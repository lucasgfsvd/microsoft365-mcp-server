import { NO_SENDING, parseSince, type PromptDefinition } from "./types.js";

const q = encodeURIComponent;

export function triageUrl(folder: string, since: Date): string {
  return (
    `/me/mailFolders/${q(folder)}/messages?$filter=${q(`receivedDateTime ge ${since.toISOString()}`)}` +
    `&$orderby=${q("receivedDateTime desc")}&$top=50` +
    "&$select=id,subject,from,receivedDateTime,isRead,importance,flag,hasAttachments,bodyPreview"
  );
}

export const inboxTriage: PromptDefinition = {
  name: "inbox-triage",
  title: "Triage my inbox",
  description: "Sort recent mail into act / reply / read / ignore, and draft the replies worth writing.",
  arguments: [
    { name: "since", description: 'How far back: "24h" (default), "3d", or a date such as 2026-09-01.' },
    { name: "folder", description: 'Mail folder: a well-known name such as "inbox" (default) or a folder id.' },
  ],
  requiredTools: ["mail_get_message"],
  render(args, ctx) {
    const since = parseSince(args.since, ctx.now);
    const folder = args.folder || "inbox";
    const url = triageUrl(folder, since);
    const fetch = ctx.tools.has("graph_batch_get")
      ? `Fetch them in one call: graph_batch_get with a single request {"id":"mail","url":"${url}"}.`
      : ctx.tools.has("mail_list_messages")
        ? `Call mail_list_messages for folder "${folder}" and keep only messages received after ${since.toISOString()}.`
        : "No listing tool is enabled; say so and stop.";
    const drafting = ctx.tools.has("mail_create_draft")
      ? 'For each Reply item, draft a short reply with mail_create_draft (to the sender, subject prefixed "Re: "), matching the tone of the original.'
      : "Writes are disabled, so write each suggested reply in your answer instead of drafting it.";

    return [
      `Triage the "${folder}" mail folder: everything received since ${since.toISOString()} (the user's time zone is ${ctx.timeZone}).`,
      "",
      `1. ${fetch} If more than 50 come back, work on the newest 50 and say how many remain.`,
      "2. Judge each message from sender, subject and bodyPreview. Open one with mail_get_message only when the preview is not enough to decide.",
      "3. Put every message in exactly one group:",
      "   - Act now: needs the user to do something, with a deadline or a person waiting.",
      "   - Reply: needs an answer but nothing else.",
      "   - Read later: worth reading, no action.",
      "   - Ignore: newsletters, notifications, FYIs.",
      "4. Present the groups in that order. One line per message: sender, subject, and why it is in that group. Mention deadlines explicitly.",
      `5. ${drafting}`,
      "",
      NO_SENDING,
    ].join("\n");
  },
};
