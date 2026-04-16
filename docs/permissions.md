# Microsoft Graph permissions

Request only the scopes you actually need — the server degrades gracefully when a tool's scopes aren't granted.

## Delegated permissions (device-code, interactive)

| Surface | Read | Write |
|---|---|---|
| Mail | `Mail.Read` | `Mail.Send`, `Mail.ReadWrite` |
| Calendar | `Calendars.Read`, `Calendars.Read.Shared` | `Calendars.ReadWrite` |
| Contacts / People | `Contacts.Read`, `People.Read` | `Contacts.ReadWrite` |
| Files / OneDrive / SharePoint | `Files.Read.All`, `Sites.Read.All` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` |
| Teams (channels) | `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All` | `ChannelMessage.Send` |
| Teams (chats) | `Chat.Read` | `ChatMessage.Send` |
| Tasks (To Do + Planner) | `Tasks.Read` | `Tasks.ReadWrite` |
| OneNote | `Notes.Read` | `Notes.ReadWrite` |
| Signed-in user | `User.Read` | – |
| Token refresh | `offline_access` | – |

## Application permissions (client-credentials)

Use the `.All` variants of the same names — e.g. `Mail.Read` → `Mail.Read`, `Files.ReadWrite.All` → `Files.ReadWrite.All`. Scopes sent at runtime always collapse to the single `https://graph.microsoft.com/.default` — the actual permissions are whatever was consented in the app registration.

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
