import { NO_SENDING, type PromptDefinition } from "./types.js";

export const meetingPrep: PromptDefinition = {
  name: "meeting-prep",
  title: "Prepare for a meeting",
  description: "Gather attendees, recent mail, related files and past meetings into a one-page brief.",
  arguments: [
    {
      name: "meeting",
      description: "Event id, or words from its subject. Default: the next meeting on the calendar.",
    },
  ],
  requiredTools: ["calendar_list_events", "calendar_get_event"],
  render(args, ctx) {
    const start = ctx.now.toISOString();
    const end = new Date(ctx.now.getTime() + 7 * 86_400_000).toISOString();
    const which = args.meeting?.trim();
    // Graph event ids are long base64 strings; anything else is a subject search.
    const find =
      which && /^[A-Za-z0-9+/=_-]{60,}$/.test(which)
        ? `Get the event with calendar_get_event, id "${which}".`
        : `Call calendar_list_events with startDateTime "${start}" and endDateTime "${end}", then pick ` +
          (which
            ? `the first event whose subject contains "${which}" (case-insensitive).`
            : "the next one that is not cancelled, not all-day and not marked free.") +
          " If none match, say so and stop.";

    const gather: string[] = [];
    if (ctx.tools.has("graph_search")) {
      gather.push(
        'Recent mail about it: graph_search with entityTypes ["message"] and a query built from the subject\'s distinctive words.',
        'Related documents: a separate graph_search with entityTypes ["driveItem"] and the same query. (Graph refuses message, event and file types in one search, so these must be separate calls.)',
        'Earlier meetings in the series: a separate graph_search with entityTypes ["event"].',
      );
    } else if (ctx.tools.has("mail_search_messages")) {
      gather.push("Recent mail about it: mail_search_messages with the subject's distinctive words.");
    }
    if (ctx.tools.has("contacts_people_search")) {
      gather.push("Who the attendees are: contacts_people_search for any attendee you cannot place from their address.");
    }

    const briefStep = gather.length ? 4 : 3;
    return [
      `Prepare the user for a meeting. Now is ${start}; the user's time zone is ${ctx.timeZone}.`,
      "",
      `1. ${find}`,
      "2. From the event, note the organizer, attendees and their responses, time (in the user's zone), location or join link, and the body/agenda.",
      ...(gather.length ? ["3. Gather context, running independent calls in parallel:", ...gather.map((g) => `   - ${g}`)] : []),
      `${briefStep}. Write a one-page brief:`,
      "   - When and where, and how long until it starts.",
      "   - Purpose, in one or two sentences.",
      "   - Who is attending, and what each is likely to care about.",
      "   - Open threads: unanswered questions or decisions pending in the recent mail.",
      "   - Documents to have open, with links.",
      "   - Three questions or points the user should raise.",
      "Keep it scannable. Say plainly when a section has nothing, rather than padding it.",
      "",
      NO_SENDING,
    ].join("\n");
  },
};
