import { describe, it, expect } from "vitest";
import { mailTools } from "../src/tools/mail/index.js";
import { calendarTools } from "../src/tools/calendar/index.js";
import { contactsTools } from "../src/tools/contacts/index.js";
import { filesTools } from "../src/tools/files/index.js";
import { teamsTools } from "../src/tools/teams/index.js";
import { tasksTools } from "../src/tools/tasks/index.js";
import { onenoteTools } from "../src/tools/onenote/index.js";
import { callTool, findTool, makeContext, makeMockGraph } from "./helpers/mockGraph.js";
import { fetchPage } from "../src/graph/pagination.js";

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("integration: mail", () => {
  it("mail_list_messages GETs the right path with select + orderby", async () => {
    const ctx = makeContext();
    await callTool(findTool(mailTools, "mail_list_messages"), ctx, { folder: "Inbox" });
    expect(ctx.mock.calls).toHaveLength(1);
    const call = ctx.mock.calls[0]!;
    expect(call.method).toBe("GET");
    expect(call.path).toBe("/me/mailFolders/Inbox/messages");
    expect(call.query.orderby).toBe("receivedDateTime desc");
    expect(String(call.query.select)).toContain("subject");
  });

  it("mail_list_messages adds isRead filter when unreadOnly=true", async () => {
    const ctx = makeContext();
    await callTool(findTool(mailTools, "mail_list_messages"), ctx, {
      folder: "Inbox",
      unreadOnly: true,
    });
    expect(ctx.mock.calls[0]!.query.filter).toBe("isRead eq false");
  });

  it("mail_send_message POSTs to /me/sendMail with the message envelope", async () => {
    const ctx = makeContext();
    await callTool(findTool(mailTools, "mail_send_message"), ctx, {
      to: [{ address: "alice@example.com", name: "Alice" }],
      subject: "Hi",
      body: "Hello",
    });
    expect(ctx.mock.calls).toHaveLength(1);
    const call = ctx.mock.calls[0]!;
    expect(call.method).toBe("POST");
    expect(call.path).toBe("/me/sendMail");
    const body = call.body as { message: { subject: string; toRecipients: Array<{ emailAddress: { address: string } }> }; saveToSentItems: boolean };
    expect(body.message.subject).toBe("Hi");
    expect(body.message.toRecipients[0]!.emailAddress.address).toBe("alice@example.com");
    expect(body.saveToSentItems).toBe(true);
  });
});

describe("integration: calendar", () => {
  it("calendar_list_events without window calls /me/events", async () => {
    const ctx = makeContext();
    await callTool(findTool(calendarTools, "calendar_list_events"), ctx, {});
    expect(ctx.mock.calls[0]!.path).toBe("/me/events");
  });

  it("calendar_list_events with start+end calls /me/calendarView", async () => {
    const ctx = makeContext();
    await callTool(findTool(calendarTools, "calendar_list_events"), ctx, {
      startDateTime: "2026-04-20T00:00:00",
      endDateTime: "2026-04-21T00:00:00",
    });
    const path = ctx.mock.calls[0]!.path;
    expect(path).toContain("/me/calendarView");
    expect(path).toContain("startDateTime=2026-04-20T00%3A00%3A00");
    expect(path).toContain("endDateTime=2026-04-21T00%3A00%3A00");
  });

  it("calendar_create_event POSTs the event payload with attendees", async () => {
    const ctx = makeContext();
    await callTool(findTool(calendarTools, "calendar_create_event"), ctx, {
      subject: "Sync",
      start: { dateTime: "2026-05-01T10:00:00", timeZone: "UTC" },
      end: { dateTime: "2026-05-01T10:30:00", timeZone: "UTC" },
      attendees: [{ email: "bob@example.com" }],
    });
    const call = ctx.mock.calls[0]!;
    expect(call.method).toBe("POST");
    expect(call.path).toBe("/me/events");
    const body = call.body as { subject: string; attendees: Array<{ type: string; emailAddress: { address: string } }> };
    expect(body.subject).toBe("Sync");
    expect(body.attendees[0]!.emailAddress.address).toBe("bob@example.com");
    expect(body.attendees[0]!.type).toBe("required");
  });
});

describe("integration: contacts", () => {
  it("contacts_list GETs /me/contacts", async () => {
    const ctx = makeContext();
    await callTool(findTool(contactsTools, "contacts_list"), ctx, {});
    expect(ctx.mock.calls[0]!.path).toBe("/me/contacts");
  });

  it("contacts_create POSTs to /me/contacts with the body", async () => {
    const ctx = makeContext();
    await callTool(findTool(contactsTools, "contacts_create"), ctx, {
      givenName: "Alice",
      surname: "Smith",
      emailAddresses: [{ address: "alice@example.com" }],
    });
    const call = ctx.mock.calls[0]!;
    expect(call.method).toBe("POST");
    expect(call.path).toBe("/me/contacts");
    expect((call.body as { givenName: string }).givenName).toBe("Alice");
  });
});

describe("integration: files", () => {
  it("files_list_children defaults to /me/drive/root/children", async () => {
    const ctx = makeContext();
    await callTool(findTool(filesTools, "files_list_children"), ctx, {});
    expect(ctx.mock.calls[0]!.path).toBe("/me/drive/root/children");
  });

  it("files_list_children with siteId targets /sites/<id>/drive", async () => {
    const ctx = makeContext();
    await callTool(findTool(filesTools, "files_list_children"), ctx, { siteId: "site123" });
    expect(ctx.mock.calls[0]!.path).toBe("/sites/site123/drive/root/children");
  });

  it("files_list_children with path uses the root: syntax", async () => {
    const ctx = makeContext();
    await callTool(findTool(filesTools, "files_list_children"), ctx, { path: "/Reports/2026" });
    expect(ctx.mock.calls[0]!.path).toBe("/me/drive/root:/Reports/2026:/children");
  });

  it("files_create_folder POSTs the folder spec with conflictBehavior=rename", async () => {
    const ctx = makeContext();
    await callTool(findTool(filesTools, "files_create_folder"), ctx, {
      parentPath: "/Projects",
      name: "Acme",
    });
    const call = ctx.mock.calls[0]!;
    expect(call.method).toBe("POST");
    expect(call.path).toBe("/me/drive/root:/Projects:/children");
    const body = call.body as { name: string; folder: object; "@microsoft.graph.conflictBehavior": string };
    expect(body.name).toBe("Acme");
    expect(body["@microsoft.graph.conflictBehavior"]).toBe("rename");
  });

  it("files_create_folder rejects path with .. segments via DrivePath schema", async () => {
    const tool = findTool(filesTools, "files_create_folder");
    expect(() => tool.inputSchema.parse({ parentPath: "/Projects/../Etc", name: "Acme" })).toThrow(/\.\./);
  });
});

describe("integration: teams", () => {
  it("teams_list_joined GETs /me/joinedTeams", async () => {
    const ctx = makeContext();
    await callTool(findTool(teamsTools, "teams_list_joined"), ctx, {});
    expect(ctx.mock.calls[0]!.path).toBe("/me/joinedTeams");
  });

  it("teams_post_channel_message POSTs the body envelope", async () => {
    const ctx = makeContext();
    await callTool(findTool(teamsTools, "teams_post_channel_message"), ctx, {
      teamId: "T1",
      channelId: "C1",
      body: "hello",
      contentType: "text",
    });
    const call = ctx.mock.calls[0]!;
    expect(call.method).toBe("POST");
    expect(call.path).toBe("/teams/T1/channels/C1/messages");
    const body = call.body as { body: { contentType: string; content: string } };
    expect(body.body.content).toBe("hello");
    expect(body.body.contentType).toBe("text");
  });
});

describe("integration: tasks", () => {
  it("todo_list_lists GETs /me/todo/lists", async () => {
    const ctx = makeContext();
    await callTool(findTool(tasksTools, "todo_list_lists"), ctx, {});
    expect(ctx.mock.calls[0]!.path).toBe("/me/todo/lists");
  });

  it("todo_create_task wraps body + dueDateTime in the Graph envelope", async () => {
    const ctx = makeContext();
    await callTool(findTool(tasksTools, "todo_create_task"), ctx, {
      listId: "L1",
      title: "Review PR",
      body: "details",
      dueDateTime: "2026-05-10T17:00:00",
      importance: "high",
    });
    const call = ctx.mock.calls[0]!;
    expect(call.method).toBe("POST");
    expect(call.path).toBe("/me/todo/lists/L1/tasks");
    const body = call.body as { title: string; body: { content: string }; dueDateTime: { dateTime: string; timeZone: string }; importance: string };
    expect(body.title).toBe("Review PR");
    expect(body.body.content).toBe("details");
    expect(body.dueDateTime.dateTime).toBe("2026-05-10T17:00:00");
    expect(body.dueDateTime.timeZone).toBe("UTC");
    expect(body.importance).toBe("high");
  });
});

describe("integration: onenote", () => {
  it("onenote_list_notebooks GETs /me/onenote/notebooks", async () => {
    const ctx = makeContext();
    await callTool(findTool(onenoteTools, "onenote_list_notebooks"), ctx, {});
    expect(ctx.mock.calls[0]!.path).toBe("/me/onenote/notebooks");
  });

  it("onenote_create_page sends xhtml+xml with injected <title>", async () => {
    const ctx = makeContext();
    await callTool(findTool(onenoteTools, "onenote_create_page"), ctx, {
      sectionId: "S1",
      title: "Weekly review",
      html: "<div>body</div>",
    });
    const call = ctx.mock.calls[0]!;
    expect(call.method).toBe("POST");
    expect(call.path).toBe("/me/onenote/sections/S1/pages");
    expect(call.headers["Content-Type"]).toBe("application/xhtml+xml");
    expect(call.body).toContain("<title>Weekly review</title>");
    expect(call.body).toContain("<div>body</div>");
  });
});

// ---------------------------------------------------------------------------
// Pagination helper — exercises the nextLink branch and search escaping.
// ---------------------------------------------------------------------------

describe("integration: pagination", () => {
  it("fetchPage chains top + filter + select onto the request", async () => {
    const { graph, calls } = makeMockGraph();
    await fetchPage(graph, "/me/messages", {
      top: 25,
      filter: "isRead eq false",
      select: ["id", "subject"],
    });
    const call = calls[0]!;
    expect(call.path).toBe("/me/messages");
    expect(call.query.top).toBe(25);
    expect(call.query.filter).toBe("isRead eq false");
    expect(call.query.select).toBe("id,subject");
  });

  it("fetchPage with nextLink calls the opaque cursor URL directly", async () => {
    const { graph, calls } = makeMockGraph();
    const nextLink = "https://graph.microsoft.com/v1.0/me/messages?$skiptoken=abc";
    await fetchPage(graph, "/me/messages", { nextLink });
    expect(calls).toHaveLength(1);
    expect(calls[0]!.path).toBe(nextLink);
    // No query chaining when using nextLink.
    expect(calls[0]!.query).toEqual({});
  });

  it("fetchPage escapes embedded quotes in the search term", async () => {
    const { graph, calls } = makeMockGraph();
    await fetchPage(graph, "/me/messages", { search: 'subject "quoted"' });
    expect(calls[0]!.query.search).toBe('"subject \\"quoted\\""');
  });
});
