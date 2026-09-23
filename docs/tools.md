# Tool reference

The full catalogue of tools exposed by the server, grouped by Microsoft 365 surface. Tools marked ✏️ are mutating and are hidden from `tools/list` until writes are enabled (see [permissions.md](./permissions.md) and the README's "Configuration" section).

Every tool name is stable across releases — the catalogue grows additively. If a name needs to change we'll keep the old name as an alias for one minor version with a deprecation log.

---

## 🔐 Auth — 2 tools

Start here: every other tool needs a signed-in session. Neither of these is mutating, so both remain available when writes are disabled — otherwise a read-only deployment could never authenticate.

| Tool | Scopes | Mutating |
|---|---|:---:|
| `auth_status` | — | |
| `auth_sign_in` | — | |

- `auth_sign_in` starts the interactive flow and returns the verification URL and device code **immediately**, rather than blocking for the minutes the browser step takes. The flow finishes in the background.
- `auth_status` reports `signed-in` / `sign-in-required`, and surfaces a device code that is still waiting to be entered.
- The server never starts a sign-in on its own. Launching it issues no prompt; a code appears only when you ask for one.

## ⚡ Graph — 3 tools

| Tool | Scopes | Mutating |
|---|---|:---:|
| `graph_batch_get` | *(whatever the sub-requests need)* | |
| `graph_search` | `Mail.Read`, `Files.Read.All`, `Sites.Read.All` | |
| `graph_delta` | `Mail.Read`, `Calendars.Read`, `Files.Read`, `Contacts.Read`, `Tasks.Read` (per resource) | |

Runs up to 20 Graph **GET** requests as a single round trip. Measured at ~2.4× faster than four separate tool calls against a live tenant, and the gap widens with more requests since it stays one round trip either way.

Each item takes an `id` you choose and a Graph-relative `url`; results come back in the order requested, each with its own `status`, so one failing sub-request does not lose the others.

**GET only, deliberately.** Batching arbitrary methods would route mutations around `writeGuard`, which is the single place writes are authorised. Sub-request URLs are validated to be Graph-relative — an absolute URL would send your token to a host of the caller's choosing.

`graph_search` runs one relevance-ranked query across mail, files, SharePoint, Teams or people — use it when you don't already know which surface something lives on, rather than running four `*_search` tools and merging by hand.

Graph combines entity types only within the SharePoint/OneDrive family (`driveItem`, `list`, `listItem`, `site`). `message`, `event`, `chatMessage` and `person` each need their own call — **including `message` and `event`, which cannot be combined**, despite both being mail-ish. The tool refuses an invalid combination before spending a round trip.

`graph_delta` is incremental sync: what was added, changed **or deleted** since last time, for mail (one folder, default inbox), calendar (a `startDateTime`–`endDateTime` window), drive, contacts, or one To Do list. The first call returns the current state and a `deltaLink`; pass that back as `cursor` to get only later changes. Deletions come back in `removed`, which is the reason to use delta at all — a listing can only show that something is absent.

- **The server stores nothing.** The `deltaLink` goes back to the caller, so there is no sync state to go stale or migrate.
- **Large first syncs are cut at `maxItems`** (default 500) with `complete: false` and a `nextLink` to continue. Treat the sync as done only once a `deltaLink` arrives. Drive pages can overshoot the cap slightly, and drive may repeat an item within one sync; the last occurrence wins.
- **Cursors must be Graph delta URLs.** The Graph client attaches your token to whatever URL it is given, so anything else is refused before a request is made.

## 📧 Mail (Outlook) — 9 tools

| Tool | Scopes | Mutating |
|---|---|:---:|
| `mail_list_messages` | `Mail.Read` | |
| `mail_search_messages` | `Mail.Read` | |
| `mail_get_message` | `Mail.Read` | |
| `mail_list_folders` | `Mail.Read` | |
| `mail_list_attachments` | `Mail.Read` | |
| `mail_send_message` | `Mail.Send` | ✏️ |
| `mail_create_draft` | `Mail.ReadWrite` | ✏️ |
| `mail_reply_message` | `Mail.Send` | ✏️ |
| `mail_delete_message` | `Mail.ReadWrite` | ✏️ |

## 📅 Calendar — 8 tools

| Tool | Scopes | Mutating |
|---|---|:---:|
| `calendar_list_events` | `Calendars.Read` | |
| `calendar_get_event` | `Calendars.Read` | |
| `calendar_list_calendars` | `Calendars.Read` | |
| `calendar_find_meeting_times` | `Calendars.Read.Shared`, `Calendars.Read` | |
| `calendar_get_free_busy` | `Calendars.Read.Shared` | |
| `calendar_create_event` | `Calendars.ReadWrite` | ✏️ |
| `calendar_update_event` | `Calendars.ReadWrite` | ✏️ |
| `calendar_delete_event` | `Calendars.ReadWrite` | ✏️ |

## 👥 Contacts & People — 6 tools

| Tool | Scopes | Mutating |
|---|---|:---:|
| `contacts_list` | `Contacts.Read` | |
| `contacts_search` | `Contacts.Read` | |
| `contacts_people_search` | `People.Read` | |
| `contacts_create` | `Contacts.ReadWrite` | ✏️ |
| `contacts_update` | `Contacts.ReadWrite` | ✏️ |
| `contacts_delete` | `Contacts.ReadWrite` | ✏️ |

## 📁 Files (OneDrive + SharePoint) — 11 tools

| Tool | Scopes | Mutating |
|---|---|:---:|
| `files_list_children` | `Files.Read.All`, `Sites.Read.All` | |
| `files_get_item` | `Files.Read.All`, `Sites.Read.All` | |
| `files_search` | `Files.Read.All`, `Sites.Read.All` | |
| `files_download` | `Files.Read.All`, `Sites.Read.All` | |
| `files_list_drives` | `Files.Read.All` | |
| `sites_search` | `Sites.Read.All` | |
| `files_upload` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |
| `files_create_folder` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |
| `files_delete` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |
| `files_copy` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |
| `files_share` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |

`files_upload` and `files_copy` use a Graph upload session above 4 MB, sent in 10 MiB chunks with retry and resume. `files_upload` is bounded by `MCP_MAX_MESSAGE_MB`, because its content arrives base64-encoded in the tool call: about 47 MB at the default of 64. `files_copy` downloads and re-uploads, so it holds the file in memory.

## 💬 Teams — 9 tools

| Tool | Scopes | Mutating |
|---|---|:---:|
| `teams_list_joined` | `Team.ReadBasic.All` | |
| `teams_list_channels` | `Channel.ReadBasic.All` | |
| `teams_list_channel_messages` | `ChannelMessage.Read.All` | |
| `teams_get_message_replies` | `ChannelMessage.Read.All` | |
| `teams_list_chats` | `Chat.Read` | |
| `teams_list_chat_messages` | `Chat.Read` | |
| `teams_post_channel_message` | `ChannelMessage.Send` | ✏️ |
| `teams_reply_channel_message` | `ChannelMessage.Send` | ✏️ |
| `teams_post_chat_message` | `ChatMessage.Send` | ✏️ |

## ✅ Tasks (To Do + Planner) — 8 tools

| Tool | Scopes | Mutating |
|---|---|:---:|
| `todo_list_lists` | `Tasks.Read` | |
| `todo_list_tasks` | `Tasks.Read` | |
| `todo_create_task` | `Tasks.ReadWrite` | ✏️ |
| `todo_complete_task` | `Tasks.ReadWrite` | ✏️ |
| `planner_list_plans` | `Tasks.Read` | |
| `planner_list_tasks` | `Tasks.Read` | |
| `planner_create_task` | `Tasks.ReadWrite` | ✏️ |
| `planner_complete_task` | `Tasks.ReadWrite` | ✏️ |

## 📓 OneNote — 6 tools

| Tool | Scopes | Mutating |
|---|---|:---:|
| `onenote_list_notebooks` | `Notes.Read` | |
| `onenote_list_sections` | `Notes.Read` | |
| `onenote_list_pages` | `Notes.Read` | |
| `onenote_get_page_content` | `Notes.Read` | |
| `onenote_create_page` | `Notes.ReadWrite` | ✏️ |
| `onenote_delete_page` | `Notes.ReadWrite` | ✏️ |

## 📊 Excel — 17 tools

| Tool | Scopes | Mutating |
|---|---|:---:|
| `excel_create_session` | `Files.ReadWrite.All` | |
| `excel_close_session` | `Files.ReadWrite.All` | |
| `excel_list_worksheets` | `Files.Read.All` | |
| `excel_get_range` | `Files.Read.All` | |
| `excel_list_tables` | `Files.Read.All` | |
| `excel_get_table_rows` | `Files.Read.All` | |
| `excel_update_range` | `Files.ReadWrite.All` | ✏️ |
| `excel_add_table_rows` | `Files.ReadWrite.All` | ✏️ |
| `excel_run_workbook_calculation` | `Files.ReadWrite.All` | ✏️ |
| `excel_create_workbook` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |
| `excel_create_from_template` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |
| `excel_add_worksheet` | `Files.ReadWrite.All` | ✏️ |
| `excel_delete_worksheet` | `Files.ReadWrite.All` | ✏️ |
| `excel_rename_worksheet` | `Files.ReadWrite.All` | ✏️ |
| `excel_create_table` | `Files.ReadWrite.All` | ✏️ |
| `excel_set_formula` | `Files.ReadWrite.All` | ✏️ |
| `excel_clear_range` | `Files.ReadWrite.All` | ✏️ |

For multi-step edits, call `excel_create_session` first and pass the returned session id to subsequent tools — Graph requires this for write operations.

## 📝 Word — 10 tools

| Tool | Scopes | Mutating |
|---|---|:---:|
| `word_read_text` | `Files.Read.All`, `Sites.Read.All` | |
| `word_list_paragraphs` | `Files.Read.All`, `Sites.Read.All` | |
| `word_create_document` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |
| `word_create_from_template` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |
| `word_replace_text` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |
| `word_append_paragraph` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |
| `word_append_heading` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |
| `word_append_bullets` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |
| `word_insert_paragraph_at` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |
| `word_delete_paragraph` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |

`word_replace_text` matches against each `<w:t>` run's contents — text broken across runs (e.g., spell-check splits) may not match. For brand-new documents prefer `word_create_document` (uses the `docx` library to emit a valid styles.xml) or `word_create_from_template` (preserves the template's themes and styles).

## 🎞️ PowerPoint — 8 tools

| Tool | Scopes | Mutating |
|---|---|:---:|
| `powerpoint_list_slides` | `Files.Read.All`, `Sites.Read.All` | |
| `powerpoint_get_slide_text` | `Files.Read.All`, `Sites.Read.All` | |
| `powerpoint_extract_all_text` | `Files.Read.All`, `Sites.Read.All` | |
| `powerpoint_create_deck` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |
| `powerpoint_create_from_template` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |
| `powerpoint_replace_text` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |
| `powerpoint_add_slide` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |
| `powerpoint_delete_slide` | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | ✏️ |

`powerpoint_add_slide` builds a minimal title-and-body slide independent of the deck's master — visual style may differ from existing slides. For richer templating, use `powerpoint_create_from_template` or `powerpoint_create_deck` followed by `powerpoint_replace_text`.

---

## Conventions

- All `*_create_*` and `*_create_from_template` tools accept `parentPath` (a folder path under the chosen drive) and `filename`. Paths must start with `/`.
- All paginated tools accept `top`, `skip`, and `nextLink` from [`PaginationInput`](../src/util/schema.ts).
- Mutating tools throw `WriteBlockedError` if writes are not enabled — the error message names the exact env var to set (`MCP_ENABLE_WRITES` or `MCP_ENABLE_<SURFACE>_WRITE`).
- Drive scope is selected uniformly: pass `driveId` for an explicit drive, `siteId` for a SharePoint site's default drive, or omit both to target the user's OneDrive (`/me/drive`).
- Prefer `graph_batch_get` when you need several independent reads; it is one round trip instead of N.
- Graph tools called while signed out return an error pointing at `auth_sign_in`, instead of blocking on a prompt the client may not render.
- Errors are normalized to `<code>: <message>` — Graph's `code` and `message` are preserved, so you can pattern-match on the code from a client.
