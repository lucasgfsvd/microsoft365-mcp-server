# Microsoft Graph permissions

Request only the scopes you actually need — the server degrades gracefully when a tool's scopes aren't granted.

## Delegated permissions (device-code, interactive)

| Surface | Read | Write |
|---|---|---|
| Mail | `Mail.Read`, `MailboxSettings.Read` | `Mail.Send`, `Mail.ReadWrite`, `MailboxSettings.ReadWrite` |
| Calendar | `Calendars.Read`, `Calendars.Read.Shared` | `Calendars.ReadWrite` |
| Contacts / People | `Contacts.Read`, `People.Read`, `User.ReadBasic.All`, `Presence.Read.All` | `Contacts.ReadWrite` |
| Files / OneDrive / SharePoint | `Files.Read.All`, `Sites.Read.All` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` |
| Teams (channels) | `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All` | `ChannelMessage.Send` |
| Teams (chats) | `Chat.Read` | `ChatMessage.Send`, `Chat.ReadWrite` |
| Tasks (To Do + Planner) | `Tasks.Read` | `Tasks.ReadWrite` |
| OneNote | `Notes.Read` | `Notes.ReadWrite` |
| Signed-in user | `User.Read` | – |
| Token refresh | `offline_access` | – |

Every scope above can be consented to by users themselves. Adding one means each user consents once more at their next sign-in.

## Scopes only an admin can grant

`Place.Read.All` (meeting rooms) and `User.Read.All` (a colleague's job title, department and phone) need an admin's approval. Asking for an unapproved scope fails the whole sign-in, so they are not requested by default: set `MCP_ENABLE_ADMIN_SCOPES=true` once an admin has consented, and `calendar_list_rooms` appears. In app-only mode the app's own permissions decide instead.

## Shared mailboxes and delegated calendars

With `MCP_ENABLE_SHARED_MAILBOXES=true`, mail and calendar tools take a `mailbox` argument naming another mailbox the user can open: a shared mailbox, or a calendar they are a delegate of. Exchange enforces the access; the server only addresses the request to `/users/{mailbox}`. The option adds these delegated scopes, so each user consents once more at the next sign-in:

| | Read | Write |
|---|---|---|
| Mail | `Mail.Read.Shared` | `Mail.ReadWrite.Shared`, `Mail.Send.Shared` |
| Calendar | `Calendars.Read.Shared` (already requested) | `Calendars.ReadWrite.Shared` |

Where the option is off, or the tool is not a mail or calendar tool, a `mailbox` argument is refused rather than ignored, so a request can never quietly act on the user's own mailbox instead.

## Application permissions (client-credentials)

Use the application versions of the same names, e.g. `Mail.ReadWrite`, `Calendars.ReadWrite`, `Files.ReadWrite.All`, `Tasks.ReadWrite.All`. Scopes sent at runtime always collapse to the single `https://graph.microsoft.com/.default` — the actual permissions are whatever was consented in the app registration.

Application permissions reach every mailbox and drive in the tenant. The server works for one user, `MCP_USER`, and offers no way to name another, but the token itself is not limited: restrict the app with an [application access policy](https://learn.microsoft.com/graph/auth-limit-mailbox-access) (or Exchange RBAC for applications) to the mailboxes it should touch.

Not available app-only, and hidden in this mode: OneNote, posting to Teams channels and chats, and `graph_search`.

## Per-surface opt-out

Even within a granted scope you can disable specific tools:

```bash
MCP_DISABLED_TOOLS=mail_delete_message,files_delete,teams_post_channel_message
```

Or disable an entire surface's writes while leaving others on:

```bash
MCP_ENABLE_WRITES=true
MCP_DISABLED_TOOLS=mail_send_message,mail_delete_message,mail_reply_message
```
