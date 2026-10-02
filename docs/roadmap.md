# Roadmap

**`0.1.0` is published when the owner decides, after the items below.** After that release the project is not actively maintained (see the
README's *Maintenance status*), so anything wanted in the published package has
to land first. The release itself is described in [releasing.md](./releasing.md)
and is already prepared: npm scope claimed, workflow ready, MCP Registry entry
written.

Each item says why it matters and roughly how big it is (S: an afternoon,
M: a day or two, L: more, or infrastructure). Every new or changed tool gets a
live-test step in [`scripts/live/`](../scripts/live/README.md) before it counts
as done: every live pass so far has found something the unit tests could not.

---

## Before 0.1.0: company use and everyday actions

Added after a review of what a user asks for in chat, and what a company deployment needs. Built and unit-tested; each needs its live pass before it counts as done.

| | Feature | Why | Size |
|---|---|---|---|
| ☐ | **App-only mode for a named user** (`MCP_USER`, certificate credentials, delegated-only tools hidden) | Documented for unattended use, but every tool used `/me`, which Graph refuses without a signed-in user. Live test needs an app registration with admin consent (`scripts/live/apponly.mjs`) | M |
| ☐ | **Shared mailboxes and delegated calendars** (`mailbox`, `MCP_ENABLE_SHARED_MAILBOXES`) | Shared inboxes and assistants managing a calendar | S |
| ☑ | **Standard tool annotations** | Clients decide what to confirm from them | S |
| ☑ | **Progress and cancellation** for large transfers | Large uploads and downloads ran silently and could not be stopped | S |
| ☐ | **Files readable by the model**: PDFs as text, images as images, mail attachments | "Summarise the PDF Alice sent" could not work | M |
| ☐ | **Mail**: send with attachments, forward, move, mark and flag, send a draft, out-of-office | The commonest mail requests after reading and replying | M |
| ☐ | **Calendar**: answer invitations, cancel meetings, recurrence, rooms (admin-gated) | | M |
| ☐ | **People**: manager, reports, profiles, presence | | S |
| ☐ | **Files**: move and rename, PDF export, permissions, SharePoint lists, recent files | | M |
| ☐ | **Teams**: direct messages, @mentions, editing your own messages | | S |
| ☐ | **Tasks**: update To Do and Planner tasks, steps, checklists, buckets | | S |
| ☐ | **OneNote** append to a page; **Excel** charts and formatting | | S |

Not included: meeting transcripts (left out of this round), and "shared with me" (Microsoft retires the API after November 2026 with no replacement).

## Before 0.1.0 (first round, done)

| | Feature | Why | Size |
|---|---|---|---|
| ☑ | **Upload from a local file**: `files_upload` takes a `localPath` inside a configured `MCP_UPLOAD_DIR`, streamed from disk through an upload session | The mirror of `saveToDisk` downloads. Lifts the ~47 MB cap that base64 in the tool call imposes, and keeps file bytes out of the conversation. Confined to one folder so a prompt cannot make the server upload arbitrary local files | M |
| ☑ | **Reply drafts in the thread**: a `mail_create_reply_draft` tool (Graph `createReply` / `createReplyAll`) | `inbox-triage` currently drafts replies as new "Re:" messages outside the conversation, because the only reply tool sends immediately | S |
| ☑ | **Date filters on `mail_list_messages`**: `receivedAfter` / `receivedBefore` | The prompts work around their absence with raw batch queries; a model using the tool directly cannot ask for "this week's mail" | S |
| ☑ | **Delete tools for To Do and Planner tasks** | Tasks can be created but never removed through the server; the live tests leave completed tasks behind for the same reason | S |
| ☑ | **Create OneNote notebooks and sections** | An account with no notebook cannot use the OneNote tools at all (seen live); pages are the only thing that can be created today | S |
| ☑ | **Excel as text in resources**: `m365://drive/…` returns `.xlsx` as CSV per sheet | Word and PowerPoint already come back as text; an attached spreadsheet arrives as an unreadable blob | S |
| ☑ | **Pick up a sign-in made by another process**: re-read `authrecord.json` when a token request finds no account | The one remaining edge from the sign-in investigation (handover): a running server does not notice a sign-in completed by another process until restarted | S |
| ☑ | **Device-code sign-in that survives a container restart**: a file-based token cache when no keyring is available, encrypted with a key from the environment | The Docker image has no `libsecret`, so device-code users must sign in after every restart; client-credentials users are unaffected | M |

## Needs a decision first

| Feature | Decision | Size |
|---|---|---|
| **Change notifications (webhooks)** | [webhooks.md](./webhooks.md) recommends staying on `graph_delta`, and Azure Event Hubs delivery if real-time is needed. Needs: is there a consumer that acts on events, which Azure subscription owns the hub, and which resources matter | L |

## Not planned

| Item | Why not |
|---|---|
| Loop, Whiteboard, Viva surfaces | Graph does not expose them in a usable form |
| Planner plans in groups that are not teams, without `groupId` | Listing all of a user's groups needs `GroupMember.Read.All`, which requires admin consent. `planner_list_plans` covers teams and takes a `groupId` for the rest |
| Deleting Teams channel messages | Needs `ChannelMessage.ReadWrite`, which requires admin consent |
| A Docker image on GHCR | An unmaintained image accumulates base-image vulnerabilities; the `Dockerfile` stays, and CI builds and starts it |

---

## Done

Kept as a record. Details are in the README, [tools.md](./tools.md) and the handover.

- 102 tools across 12 surfaces, all exercised against a real tenant
- `graph_search`, `graph_batch_get`, `graph_delta`
- Prompt templates: `daily-brief`, `inbox-triage`, `meeting-prep`
- MCP resources: mail messages, OneDrive/SharePoint files, OneNote pages
- Large-file upload sessions; streaming downloads to `MCP_DOWNLOAD_DIR`
- Per-tool retry policy (a send is never repeated on 503/504)
- Pre-authenticated download links stripped from results; mail padding tidied
- Lazy sign-in; `--logout` clears this app's tokens; runs without an OS keyring
- Live-test scripts and a private sandbox team, plan and notebook
