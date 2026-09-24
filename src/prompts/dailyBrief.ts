import { dayWindow, NO_SENDING, parseDay, type PromptDefinition } from "./types.js";

const q = encodeURIComponent;

export function briefRequests(day: string, timeZone: string, now: Date): Array<{ id: string; url: string }> {
  const { start, end } = dayWindow(day, timeZone);
  const since = new Date(now.getTime() - 3 * 86_400_000).toISOString();
  return [
    {
      id: "agenda",
      url:
        `/me/calendarView?startDateTime=${q(start)}&endDateTime=${q(end)}&$orderby=${q("start/dateTime")}&$top=50` +
        "&$select=subject,start,end,location,isAllDay,showAs,isCancelled,organizer,onlineMeeting,webLink",
    },
    {
      // Graph wants $orderby properties to lead the $filter, in the same order.
      id: "unread",
      url:
        `/me/mailFolders/inbox/messages?$filter=${q(`receivedDateTime ge ${since} and isRead eq false`)}` +
        `&$orderby=${q("receivedDateTime desc")}&$top=25&$select=id,subject,from,receivedDateTime,importance,bodyPreview`,
    },
    {
      id: "flagged",
      url: `/me/messages?$filter=${q("flag/flagStatus eq 'flagged'")}&$top=25&$select=id,subject,from,receivedDateTime,flag`,
    },
    // To Do rejects a multi-field $select on lists ("Invalid request", seen live); the list shape is small anyway.
    { id: "todo-lists", url: "/me/todo/lists" },
  ];
}

// To Do answers "Invalid request" to any $filter or $select on tasks via $batch
// (seen live, even $select=title); only $top works, so completed tasks are
// dropped by the model instead of by the query.
export function taskListUrl(listId: string): string {
  return `/me/todo/lists/${q(listId)}/tasks?$top=100`;
}

export const dailyBrief: PromptDefinition = {
  name: "daily-brief",
  title: "Brief me on my day",
  description: "Agenda, unread and flagged mail, and open tasks for a day, gathered in one round trip.",
  arguments: [{ name: "date", description: "Day to brief, as YYYY-MM-DD. Default: today." }],
  requiredTools: ["graph_batch_get"],
  render(args, ctx) {
    const day = parseDay(args.date, ctx.now, ctx.timeZone);
    const requests = briefRequests(day, ctx.timeZone, ctx.now);
    return [
      `Brief the user on ${day} (time zone ${ctx.timeZone}).`,
      "",
      "1. Gather everything in one call: graph_batch_get with these requests:",
      "```json",
      JSON.stringify(requests, null, 2),
      "```",
      "2. From todo-lists, make one more graph_batch_get with a request per list (at most 20), url " +
        `"${taskListUrl("{id}").replace("%7Bid%7D", "{id}")}". ` +
        "Skip tasks whose status is completed. Keep those due on or before that day, plus any marked important.",
      "3. Any request whose status is not 2xx: say which part is missing and carry on with the rest.",
      "4. Write the brief:",
      "   - Agenda: each event with local start and end times (convert from UTC to the user's zone). Skip cancelled ones, point out overlaps, and list free blocks of 30 minutes or more between 08:00 and 18:00.",
      "   - Needs attention: unread mail from the last three days that looks like it needs the user, most important first, one line each.",
      "   - Flagged: flagged messages, one line each.",
      "   - Tasks: overdue first, then due that day, then important.",
      "   - Close with the single most important thing to do first, and why.",
      "",
      NO_SENDING,
    ].join("\n");
  },
};
