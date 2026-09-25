import { tidyText } from "./html.js";
import { enc, uriFor, values, type ResourceKind } from "./types.js";

interface Address {
  emailAddress?: { name?: string; address?: string };
}

const who = (a?: Address) => {
  const e = a?.emailAddress;
  if (!e) return "";
  return e.name && e.address && e.name !== e.address ? `${e.name} <${e.address}>` : (e.address ?? e.name ?? "");
};

interface Message {
  subject?: string;
  from?: Address;
  toRecipients?: Address[];
  ccRecipients?: Address[];
  receivedDateTime?: string;
  hasAttachments?: boolean;
  body?: { content?: string };
}

/** A header block and the body, the way someone would paste an email. */
export function formatMessage(m: Message): string {
  const lines = [
    `Subject: ${m.subject ?? "(no subject)"}`,
    `From: ${who(m.from)}`,
    m.toRecipients?.length ? `To: ${m.toRecipients.map(who).join(", ")}` : "",
    m.ccRecipients?.length ? `Cc: ${m.ccRecipients.map(who).join(", ")}` : "",
    m.receivedDateTime ? `Date: ${m.receivedDateTime}` : "",
    m.hasAttachments ? "Attachments: yes (see mail_list_attachments)" : "",
  ].filter(Boolean);
  return `${lines.join("\n")}\n\n${tidyText(m.body?.content ?? "")}\n`;
}

export const mailResource: ResourceKind = {
  key: "mail",
  requiresTool: "mail_get_message",
  template: {
    uriTemplate: "m365://mail/{messageId}",
    name: "mail-message",
    title: "Mail message",
    description: "An email as plain text: headers, then the body.",
    mimeType: "text/plain",
  },
  recent: {
    id: "mail",
    url: "/me/mailFolders/inbox/messages?$top=20&$select=id,subject,from,receivedDateTime&$orderby=receivedDateTime%20desc",
  },
  toEntries: (body) =>
    values(body).map((m) => ({
      uri: uriFor("mail", String(m.id)),
      name: `mail: ${(m.subject as string) || "(no subject)"}`,
      description: `From ${who(m.from as Address)}, ${m.receivedDateTime ?? ""}`.trim(),
      mimeType: "text/plain",
    })),
  async read(graph, [id], uri) {
    // Ask Outlook for the text rendering; HTML mail would cost far more context.
    const m = (await graph
      .api(`/me/messages/${enc(id!)}`)
      .header("Prefer", 'outlook.body-content-type="text"')
      .select("subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,body")
      .get()) as Message;
    return { uri, mimeType: "text/plain", text: formatMessage(m) };
  },
};
