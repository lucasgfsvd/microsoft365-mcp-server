<div align="center">

# Microsoft 365 MCP Server

**Bring your Microsoft 365 life into any AI assistant.**

Mail · Calendar · Contacts · OneDrive · SharePoint · Teams · To Do · Planner · OneNote · Excel · Word · PowerPoint — all behind a single [Model Context Protocol](https://modelcontextprotocol.io) server.

[![Status: alpha](https://img.shields.io/badge/status-alpha-orange?style=flat-square)](#status--whats-tested)
[![CI](https://img.shields.io/github/actions/workflow/status/lucasgfsvd/microsoft365-mcp-server/ci.yml?branch=main&style=flat-square)](https://github.com/lucasgfsvd/microsoft365-mcp-server/actions)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg?style=flat-square)](./LICENSE)
[![MCP](https://img.shields.io/badge/MCP-compatible-7B3FE4?style=flat-square)](https://modelcontextprotocol.io)

</div>

---

## Why this exists

- **One M365 surface for every MCP client.** Works out of the box with Claude Desktop, Cursor, Continue, Cline, Zed — anything that speaks MCP.
- **Safe by default.** All mutating tools (send mail, delete files, post to Teams) are **disabled** until you explicitly opt in. No surprise emails.
- **Individuals _and_ tenants.** Personal Microsoft accounts use a zero-config **device code** login. Companies can run it unattended with **client credentials** and admin consent.
- **Deep Office editing.** Not just Graph wrappers — edit real `.docx` and `.pptx` files in-place via OOXML, not just upload/download.

---

## Features

| Surface | Read | Write | Highlights |
|---|:---:|:---:|---|
| 📧 **Mail** (Outlook) | ✅ | ✅ | List, search (KQL), threaded replies, drafts, attachments |
| 📅 **Calendar** | ✅ | ✅ | `findMeetingTimes`, free/busy across attendees, online meetings |
| 👥 **Contacts & People** | ✅ | ✅ | Personal contacts + directory search (People API) |
| 📁 **Files** (OneDrive + SharePoint) | ✅ | ✅ | Drive & site navigation, search, up/download, sharing links |
| 💬 **Teams** | ✅ | ✅ | Channels, chats, post messages & replies |
| ✅ **Tasks** (To Do + Planner) | ✅ | ✅ | Lists, tasks, completions, assignments |
| 📓 **OneNote** | ✅ | ✅ | Notebooks, sections, HTML page CRUD |
| 📊 **Excel** | ✅ | ✅ | Create workbooks (from scratch or **from templates**), add/rename/delete sheets, tables, ranges, formulas, recalculation |
| 📝 **Word** | ✅ | ✅ | Create documents (from scratch or **from templates** with placeholder fill-in), append headings/bullets/paragraphs, insert at index, find-and-replace |
| 🎞️ **PowerPoint** | ✅ | ✅ | Create decks (from scratch or **from templates** with placeholder fill-in + extra slides), append/delete slides, find-and-replace across slides |

~92 tools total. The full reference is in [docs/tools.md](./docs/tools.md).

---

## Status — what's tested

This is an **early alpha**. Please calibrate expectations before depending on it:

- **Verified end-to-end against a real Microsoft 365 business tenant:** device-code auth (incl. the in-chat sign-in prompt fallback), token caching, write-gate toggling, `mail_list_folders`.
- **Unit tested (15 tests):** config parsing, the tool registry, the write-guard, and the full tool catalogue (see `test/`). CI runs these on every push.
- **Code-reviewed but not yet exercised against a live tenant:** the remaining ~90 tools. Graph passthroughs (mail / calendar / files / teams / tasks / OneNote / Excel-via-Graph) are low risk — if Graph accepts the payload, it works. The higher-risk area is the **OOXML surgery** in the PowerPoint `add_slide`/`delete_slide` and Word `append_heading`/`insert_paragraph_at`/`delete_paragraph` tools: they generate and splice raw XML, so edge-case deck/document layouts can produce a file that opens but rendered slightly off, or in the worst case refuses to open. If you hit that, **please [open an issue](https://github.com/lucasgfsvd/microsoft365-mcp-server/issues) and attach the input file** — it's the fastest way to harden those paths.
- **Not yet published** — no npm package, no Docker image, no GHCR release. Install today is **clone + build locally** (see the Quickstart below). npm / Docker artefacts will come after the alpha shakes out.

If you're evaluating this for anything more serious than experimentation, wait for the `0.2.0` tag — at which point more of the write-surface tools will have been exercised and the OOXML paths hardened based on early-alpha feedback.

---

## Quickstart — from a local build (the only working path today)

Until the npm package and Docker image are published, install by cloning + building. It takes about 90 seconds.

```bash
git clone https://github.com/lucasgfsvd/microsoft365-mcp-server
cd microsoft365-mcp-server
npm install
npm run build          # produces dist/index.js
```

Then add this to your MCP client config (Claude Desktop example — see [per-OS paths](#operating-system-config-paths) further down):

```json
{
  "mcpServers": {
    "microsoft365": {
      "command": "node",
      "args": ["/absolute/path/to/microsoft365-mcp-server/dist/index.js"],
      "env": {
        "MCP_AUTH_MODE": "device-code"
      }
    }
  }
}
```

On **Windows**, give the full path to `node.exe` and escape backslashes (`\\`) — Claude Desktop inherits the system PATH, not your shell's, and doesn't parse `~`/`$HOME`:

```json
{
  "mcpServers": {
    "microsoft365": {
      "command": "C:\\Program Files\\nodejs\\node.exe",
      "args": ["C:\\path\\to\\microsoft365-mcp-server\\dist\\index.js"],
      "env": {
        "MCP_AUTH_MODE": "device-code"
      }
    }
  }
}
```

Fully quit and reopen Claude Desktop (tray icon → *Quit* — closing the window alone does not reload the config). On first tool call the server emits a sign-in prompt as an MCP logging notification *and* as a tool-error fallback, so you'll see the URL + code directly in the chat; no need to tail stderr. Tokens are cached at `~/.microsoft365-mcp/tokencache.json` (chmod 600) so you only sign in once.

> **Writes are disabled by default.** To allow sending mail, creating events, posting to Teams, etc. add `"MCP_ENABLE_WRITES": "true"` to the `env` block above — or scope it: `"MCP_ENABLE_MAIL_WRITE": "true"`.

---

## Quickstart — via `npx` (coming once we publish)

> 🚧 **Not available yet.** The npm package `@microsoft365-mcp/server` hasn't been published — we're gating publish on alpha feedback. Once it's up, this section becomes the simplest install for anyone who isn't modifying the source.

```json
{
  "mcpServers": {
    "microsoft365": {
      "command": "npx",
      "args": ["-y", "@microsoft365-mcp/server"],
      "env": {
        "MCP_AUTH_MODE": "device-code"
      }
    }
  }
}
```

---

## Operating-system config paths

The Quickstart above shows Claude Desktop; other MCP clients (Cursor, Continue, Cline, Zed) have their own config files but all accept the same `mcpServers` block. For Claude Desktop specifically, the config file lives at:

| OS | Path |
|---|---|
| **Windows** | `%APPDATA%\Claude\claude_desktop_config.json` (i.e. `C:\Users\<you>\AppData\Roaming\Claude\claude_desktop_config.json`) |
| **macOS** | `~/Library/Application Support/Claude/claude_desktop_config.json` |
| **Linux** | `~/.config/Claude/claude_desktop_config.json` |

After editing, **fully quit Claude Desktop** (tray icon → *Quit*, not just close the window) and reopen. Verify the server loaded under **Settings → Developer** and inside the in-chat tools picker. If it's missing, the startup error is in `%APPDATA%\Claude\logs\mcp-server-microsoft365.log` on Windows, `~/Library/Logs/Claude/mcp-server-microsoft365.log` on macOS.

> Local MCP servers appear under **Settings → Developer** and the in-chat tools picker — **not** under **Settings → Connectors**. That section is for Anthropic-hosted remote connectors only.

---

## Quickstart — for teams / tenants (Docker + client credentials) — coming once we publish

> 🚧 **Not available yet.** No Docker image has been pushed to GHCR or anywhere else. Once the alpha stabilises we'll publish a multi-arch distroless image. Until then, run the local build on a server (or in your own container) with `MCP_AUTH_MODE=client-credentials` — the auth/config is identical; only the launch command differs.

Once published, the per-tenant deployment will look like this:

1. **Register an app** in your Azure tenant: Entra ID → App registrations → New registration.
2. **Add API permissions** (Application type, not Delegated) from Microsoft Graph — see [docs/permissions.md](./docs/permissions.md) for the full list.
3. **Grant admin consent** for your tenant.
4. **Create a client secret** under Certificates & secrets.
5. **Run the image** — example Claude Desktop / MCP client config:

```json
{
  "mcpServers": {
    "microsoft365": {
      "command": "docker",
      "args": [
        "run", "-i", "--rm",
        "-v", "m365-mcp-tokens:/data",
        "-e", "MCP_AUTH_MODE=client-credentials",
        "-e", "MCP_TENANT_ID",
        "-e", "MCP_CLIENT_ID",
        "-e", "MCP_CLIENT_SECRET",
        "-e", "MCP_ENABLE_WRITES=true",
        "ghcr.io/lucasgfsvd/microsoft365-mcp-server:latest"
      ],
      "env": {
        "MCP_TENANT_ID": "<your-tenant-guid>",
        "MCP_CLIENT_ID": "<your-app-id>",
        "MCP_CLIENT_SECRET": "<your-secret>"
      }
    }
  }
}
```

The image is multi-arch (`linux/amd64`, `linux/arm64`), built on distroless Node 20, runs as nonroot, and caches tokens in the `m365-mcp-tokens` named volume.

---

## Register your own Azure AD app

The default `MCP_CLIENT_ID` baked into the server is the public "Microsoft Graph Command Line Tools" client — it's fine for **trying the server out**, but for anything beyond that (production use, personal-brand audit logs, custom scope sets) you should register your own app. It's free and takes ~3 minutes.

### 1. Open the Entra admin center

Go to **[entra.microsoft.com](https://entra.microsoft.com)** (or [portal.azure.com](https://portal.azure.com) → search *App registrations*). Sign in with any Microsoft account — if you don't already have a work/school tenant, Microsoft will auto-create a free one for you.

### 2. Create the registration

Left sidebar → **Applications → App registrations** → **+ New registration**. Then:

| Field | Value |
|---|---|
| **Name** | `microsoft365-mcp-server` (or anything you like) |
| **Supported account types** | *Accounts in any organizational directory (Any Microsoft Entra ID tenant — Multitenant) and personal Microsoft accounts* |
| **Redirect URI** | Leave blank for now |

Click **Register**.

### 3. Copy the Application (client) ID

On the app's **Overview** page you'll see three GUIDs near the top:

```
Application (client) ID   →  xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx   ← this is what you want
Directory (tenant) ID     →  …
Object ID                 →  …
```

Copy the **Application (client) ID** — this becomes your `MCP_CLIENT_ID`.

### 4. Enable device-code / public-client flows

**Authentication** tab → **+ Add a platform** → **Mobile and desktop applications** → tick `https://login.microsoftonline.com/common/oauth2/nativeclient` → **Configure**. Then scroll down and set **Allow public client flows** = **Yes** → **Save**.

(Skip this step if you'll only ever use `client-credentials` — see step 6 for that instead.)

### 5. Add API permissions

**API permissions** → **+ Add a permission** → **Microsoft Graph** → **Delegated permissions**. Tick each scope listed in [docs/permissions.md](./docs/permissions.md) that matches the surfaces you want to expose — at minimum `User.Read`, `offline_access`, then whichever of `Mail.ReadWrite`, `Calendars.ReadWrite`, `Files.ReadWrite.All`, `Sites.ReadWrite.All`, `ChannelMessage.Send`, `Chat.Read`, `Tasks.ReadWrite`, `Notes.ReadWrite`, `People.Read`, etc. you need.

For **individual users**: no admin consent required — each user self-consents on first login.

### 6. (Tenant automation only) Add a client secret and admin consent

If you're running unattended with `MCP_AUTH_MODE=client-credentials`:

1. **Certificates & secrets** → **+ New client secret** → copy the value (it's only shown once). That's your `MCP_CLIENT_SECRET`.
2. **API permissions** → add the same permissions as step 5 but pick **Application permissions** instead of Delegated.
3. Click **Grant admin consent for \<tenant\>**. A Global Administrator must approve.
4. Copy the **Directory (tenant) ID** from the Overview page — that's your `MCP_TENANT_ID`.

### 7. Use it

Set `MCP_CLIENT_ID` (and for tenant use, `MCP_TENANT_ID` + `MCP_CLIENT_SECRET`) in your MCP client config:

```json
{
  "mcpServers": {
    "microsoft365": {
      "command": "npx",
      "args": ["-y", "@microsoft365-mcp/server"],
      "env": {
        "MCP_AUTH_MODE": "device-code",
        "MCP_CLIENT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
      }
    }
  }
}
```

Delete any existing `~/.microsoft365-mcp/tokencache.json` so the next login uses your new app, and you're done.

---

## Rolling out to a company / team

**Common misconception:** IT admins do **not** need to copy per-user client IDs. You register **one** app; everyone uses it.

There are two deployment models — pick based on *who* the tool acts as:

### Model A — Delegated (per-user sign-in) — **most teams want this**

One app registration shared by the whole org. Each user signs in **as themselves** and only sees their own mail / files / calendar. Alice sees Alice's mail, Bob sees Bob's — even though they share the same `MCP_CLIENT_ID`.

**IT admin does — once:**

1. Follow steps 1–5 of [Register your own Azure AD app](#register-your-own-azure-ad-app) in the company tenant.
2. On the API permissions page, click **"Grant admin consent for \<tenant\>"**. This pre-approves the scopes for the whole org so users never see a consent popup.
3. Distribute `MCP_CLIENT_ID` and `MCP_TENANT_ID` via whatever channel you already use: Intune policy, MDM, internal wiki, a shared Claude Desktop config template, a pre-baked dev image. Those two values are not secrets.

**Each user does — once:**

1. Add the config block to their MCP client with the two values IT provided.
2. On first tool call, complete the device-code sign-in **as themselves**. Their personal refresh token is cached in their own home directory.

That's it — no per-user Entra config, no IDs to collect, no IT ticket per person. Add a new hire by onboarding them to the MCP client config; they sign in and they're live.

```json
{
  "mcpServers": {
    "microsoft365": {
      "command": "npx",
      "args": ["-y", "@microsoft365-mcp/server"],
      "env": {
        "MCP_AUTH_MODE": "device-code",
        "MCP_TENANT_ID": "<company-tenant-guid>",
        "MCP_CLIENT_ID": "<company-app-id>",
        "MCP_ENABLE_WRITES": "true"
      }
    }
  }
}
```

### Model B — Application (unattended automation)

Used when the server runs on a server, cron job, or bot with **no human** behind it. The app acts as itself, not any user, and has tenant-wide reach (can read any mailbox, any file — by default).

- One app registration with **Application** permissions (not Delegated) + admin consent + a client secret.
- **Always narrow the blast radius.** Scope the app to a subset of mailboxes / sites so a compromise or prompt-injection attack can't exfiltrate the whole tenant:
  - Mail: [Application Access Policies](https://learn.microsoft.com/graph/auth-limit-mailbox-access)
  - Exchange Online more broadly: [RBAC for Applications](https://learn.microsoft.com/exchange/permissions-exo/application-rbac)
  - SharePoint: [Sites.Selected](https://learn.microsoft.com/sharepoint/dev/solution-guidance/security-apponly-azureacs) instead of `Sites.ReadWrite.All`
- Deploy the secret via your existing secret manager (Key Vault, GitHub secrets, Vault, etc.) — never into per-user configs.
- Prefer [Workload Identity Federation](https://learn.microsoft.com/entra/workload-id/workload-identity-federation) over a static client secret when the workload runs on GitHub Actions, AKS, or another federated platform.

### Which to pick?

| Question | Model A (delegated) | Model B (application) |
|---|:---:|:---:|
| Employees use this interactively from their AI client? | ✅ | |
| Runs unattended on a server / bot? | | ✅ |
| Each user should only see their own data? | ✅ | |
| Needs tenant-wide reach (e.g. compliance bot)? | | ✅ |
| Requires client secret management? | ❌ | ✅ |
| Works for personal Microsoft accounts? | ✅ | ❌ |

Most companies run **Model A for humans** and, if needed, **a separate Model B app** for automation — never the same app for both.

---

## Example prompts

Once connected, try these in your MCP client. Everything in the **Read-only** columns works with the default config. Examples marked ✏️ require `MCP_ENABLE_WRITES=true` (or the matching per-surface flag) plus the write scopes.

### Mail (Outlook)

**Read-only**
- *"Summarize my unread emails from the last 48 hours and tell me which ones look urgent."*
- *"Find every message from Alice at Acme in the last month that mentions 'contract' and list the subjects."*
- *"Open the latest email from finance and extract any attachments to a summary."*
- *"What's the oldest unanswered email in my inbox from an external sender?"*
- *"Group my unread mail by sender domain and show me the top 10."*

**✏️ Write**
- *"Draft a reply to the last email from Bob saying I'll have the numbers by Friday — don't send yet."*
- *"Send a polite chase-up to everyone who hasn't replied to the 'Q3 kickoff' thread."*
- *"Delete every promotional email from the last 30 days in the Promotions folder."*

### Calendar

**Read-only**
- *"What's on my calendar tomorrow and where are the conflicts?"*
- *"Find a 30-minute slot next Tuesday afternoon that works for alice@acme.com and bob@acme.com."*
- *"Show me the free/busy for the whole team (alice@, bob@, carol@) on Thursday between 9am and 5pm."*
- *"How many hours of meetings am I in this week, and how much is back-to-back?"*

**✏️ Write**
- *"Book that Tuesday slot as 'Q3 kickoff', 30 min, invite alice@ and bob@, add a Teams link."*
- *"Move my 3pm today with Dave to the same time tomorrow."*
- *"Cancel every internal meeting on Friday and send a note saying I'm out sick."*

### Contacts & People

- *"Find the best email for 'Priya at Acme' — check my contacts first, then the directory."*
- *"List everyone in my contacts I haven't emailed in the last 6 months."*
- ✏️ *"Add a contact for John Smith, john@example.com, tagged 'vendor'."*

### Files (OneDrive + SharePoint)

**Read-only**
- *"Search SharePoint for docs called 'architecture' edited this month and list them with owners."*
- *"What are the 20 largest files in my OneDrive?"*
- *"Find every `.docx` in the `/Proposals` folder shared externally."*
- *"Download `research-notes.pdf` from the team site and give me a summary."*

**✏️ Write**
- *"Create a folder `2026-Q2` under `/Projects/Acme` and upload the attached `kickoff.md` into it."*
- *"Generate a read-only share link for `pricing-v3.xlsx` that expires in 7 days and copy it to my clipboard."*
- *"Delete every file in `/tmp-uploads` older than 30 days."*

### Teams

**Read-only**
- *"What have I missed in #engineering over the last 24 hours? Summarize by thread."*
- *"List every 1:1 chat I have with someone outside my team."*
- *"Find the message in #incidents where Carol posted the runbook link last week."*

**✏️ Write**
- *"Post today's deploy summary to #engineering on the Platform team."*
- *"Reply to the top thread in #standup with my update: shipped X, blocked on Y."*
- *"DM Alice the calendar slot we just booked."*

### Tasks (To Do + Planner)

- *"List my overdue To Do tasks and mark the three from the 'Inbox' list that mention 'demo' as complete."*
- *"Show every Planner task assigned to me across all plans, sorted by due date."*
- ✏️ *"Create a To Do task 'Review PR #482' due Thursday, high importance, in the 'Work' list."*
- ✏️ *"In the Planner plan 'Launch Q3', create a task 'Draft press release' assigned to me, due next Monday."*

### OneNote

- *"List my notebooks and show the 5 most recently edited pages across all of them."*
- *"Get the full text of the page 'Customer conversations — April' in the 'Weekly reviews' section."*
- ✏️ *"Create a new page in the 'Weekly reviews' section summarizing this week's top 5 customer conversations."*

### Working with templates — what "template" means here

The `*_create_from_template` tools (and `files_copy`) work on **files that exist in your OneDrive or SharePoint** — your org's templates, your own `.pptx` / `.docx` / `.xlsx` in a `/Templates/` folder, a shared Teams files tab, anything you can reach via Graph.

They do **not** have access to Microsoft's built-in "File → New" gallery (*Ion Boardroom*, *Berlin*, *Modern proposal*, etc.). That gallery is served by a separate Office service (`templates.office.com`) that isn't exposed through Microsoft Graph. If you want to use one of those stock templates, the one-time workaround is: open PowerPoint/Word → *File → New* → pick the template → **Save a copy to OneDrive** (e.g. `/Templates/Ion-Boardroom.pptx`). After that it's just a file in your drive — reference it by path or by name like any other template.

Agents can also resolve a template by **name** without knowing the exact path: chain `files_search` ("find a .pptx with 'proposal' in the name") with `powerpoint_create_from_template` using the returned `id`. That's how most of the template prompts below are phrased.

### Excel (create workbooks + live editing)

- *"Open `sales-2026.xlsx` in OneDrive, read `Pipeline!A1:F50`, and tell me which deals have slipped."*
- *"List every worksheet and table in `budget.xlsx`."*
- ✏️ *"Create a new workbook `lead-tracker.xlsx` in `/Spreadsheets/` with a sheet 'Leads' seeded with headers: Name, Company, Stage, ARR, Owner."*
- ✏️ *"From template `/Templates/quarterly-report.xlsx`, create `Q2-2026-report.xlsx` in `/Reports/`, then fill `Summary!B2:B5` with [120000, 145000, 98000, 162000]."*
- ✏️ *"Search my OneDrive for a workbook with 'budget' in the name, use it as a template to create `budget-2026-Q2.xlsx` in `/Reports/`, and tell me which file you used."* (chains `files_search` + `excel_create_from_template`)
- ✏️ *"Add a 'Forecast' worksheet to `lead-tracker.xlsx`, write `=SUMIF(Leads!C:C,"Won",Leads!D:D)` into A1, then recalc."*
- ✏️ *"In `lead-tracker.xlsx` convert the range `Leads!A1:E100` into a table named `Leads` with headers."*
- ✏️ *"In `forecast.xlsx`, add a row to the `Leads` table: name='Globex', stage='Qualified', arr=24000."*
- ✏️ *"In `model.xlsx`, write 1.07 into `Assumptions!B12`, set `Summary!B2` to formula `=Assumptions!B12 * Inputs!C4`, then recalc and read `Summary!A1:D20`."*
- ✏️ *"In `stale.xlsx`, clear the contents (keep formatting) of `Sheet1!A1:Z100`, then rename 'Sheet1' to 'Archive-2025'."*

### Word (create + structured editing)

- *"Extract the full text of `proposal-v2.docx` and tell me the outline."*
- *"List every paragraph in `policy.docx` that mentions 'Section 4'."*
- ✏️ *"Create a new document `weekly-report-2026-04-16.docx` in `/Reports/` with a title 'Weekly Report', heading1 'Highlights', three bullets summarising my top accomplishments, heading1 'Blockers', two bullets, and a closing paragraph."*
- ✏️ *"From template `/Templates/proposal.docx`, create `Acme-proposal-2026-04.docx` in `/Proposals/Acme/`, replacing `{{CLIENT_NAME}}` with 'Acme Corp', `{{EFFECTIVE_DATE}}` with '2026-05-01', `{{PRICE}}` with '€24,000', and append a heading2 'Appendix A — References' followed by 3 bullets of past work."*
- ✏️ *"Find a `.docx` in my OneDrive whose name contains 'proposal', use it as a template to create `Globex-proposal.docx` in `/Proposals/Globex/`, replacing `{{CLIENT_NAME}}` with 'Globex Inc', and tell me which template file you picked."* (chains `files_search` + `word_create_from_template`)
- ✏️ *"In `changelog.docx`, append a heading2 'Release 2026-04-16' followed by a bullet list of the 5 commits I made this week."*
- ✏️ *"In `contract-template.docx`, replace every `{{CLIENT_NAME}}` with 'Acme Corp' and `{{EFFECTIVE_DATE}}` with '2026-05-01'."*
- ✏️ *"Insert a paragraph after paragraph 3 in `onboarding.docx` with the new policy text, and delete paragraph 12 (the outdated section)."*
- ✏️ *"Append a new paragraph to `changelog.docx` with today's date and the text of my last commit."*

### PowerPoint (create decks + in-place editing)

- *"List the slide titles in `Q3-review.pptx` and give me the text of slide 7."*
- *"Extract every bullet from `all-hands-april.pptx` and turn it into meeting notes."*
- ✏️ *"Create a new deck `kickoff-acme.pptx` in `/Decks/` with 4 slides: (1) title 'Acme Kickoff' with presenter notes; (2) title 'Agenda' with bullets 'Introductions, Timeline, Risks, Q&A'; (3) title 'Timeline' with bullets for each phase; (4) title 'Questions?'."*
- ✏️ *"From template `/Templates/client-kickoff.pptx`, create `kickoff-acme.pptx` in `/Decks/Acme/`, replacing `{{CLIENT}}` with 'Acme Corp', `{{DATE}}` with '2026-04-16', `{{OWNER}}` with 'Dana', and append one extra slide titled 'Our team' with bullets for each teammate."*
- ✏️ *"Find a PowerPoint template in my OneDrive whose name contains 'kickoff'. Use it to create `kickoff-globex.pptx` in `/Decks/Globex/`, replacing any `{{CLIENT}}` with 'Globex Inc' and `{{DATE}}` with today. Tell me which file you ended up using."* (chains `files_search` + `powerpoint_create_from_template`)
- ✏️ *"Append a slide titled 'Next steps' with 3 bullets to the existing deck `Q3-review.pptx`."*
- ✏️ *"Delete slide 12 from `board-deck.pptx` (it's outdated)."*
- ✏️ *"In `Q3-review.pptx`, replace every `FY24` with `FY25` and list which slides changed."*
- ✏️ *"Find the slide titled 'Risks' in `board-deck.pptx` and replace the body text with the three risks I just listed."*

### Cross-surface workflows — where an MCP actually shines

These stitch multiple tools together in a single prompt, which is the real pay-off over using Outlook/Teams/SharePoint individually:

- 🧵 **Triage & respond.** *"Read my unread mail from the last 24h, flag anything from a customer, file the rest into the right folder, and draft replies to the customer ones in my voice."*
- 📣 **Meeting prep.** *"For my 2pm with Acme: pull the last 5 emails from anyone @acme.com, summarize the Teams chat with their account team, fetch the latest version of `Acme-proposal.docx` from SharePoint, and put it all in a new OneNote page under 'Prep'."*
- 📊 **Weekly report.** *"Read `sales-2026.xlsx!Pipeline`, list the deals that moved stage this week, post the summary to #sales on Teams, and create a Planner task for me to follow up on each 'Stalled' one."*
- 🚨 **Incident follow-up.** *"From the #incidents channel, find the last 'sev2' thread, extract timeline and action items, create a OneNote page with the post-mortem template filled in, and assign each action item as a Planner task to the person who owns it."*
- 📝 **Proposal automation.** *"Take `contract-template.docx`, replace the client placeholders with 'Globex Inc', save a copy as `Globex-contract-2026-04.docx` in `/Proposals/Globex`, share it read-only with their procurement contact, and email them the link."*
- 📅 **Travel day builder.** *"Find a 90-minute gap next Wednesday, block it as 'Focus — board deck', and create a To Do list 'Board prep' with one task per slide in `board-deck.pptx`."*
- 🔍 **Cross-mailbox research.** *"Search my mail, Teams DMs, and OneNote for anything mentioning 'Project Condor' in the last 90 days, and summarize what's happening."*
- 🧹 **Cleanup.** *"Find every SharePoint file I 'own' that hasn't been opened in 12 months, list them, and for the ones under 1MB create a shared `Archive-2026` folder and move them there."*

> Many of the write-heavy prompts above will trigger multiple tool calls in sequence. If writes are disabled, the assistant will tell you exactly which scope or which `MCP_ENABLE_*` flag is missing — no silent failures.

---

## Tool reference

<details>
<summary><strong>📧 Mail</strong> — 9 tools</summary>

| Tool | Scopes | Writes |
|---|---|:---:|
| `mail_list_messages` | Mail.Read | |
| `mail_search_messages` | Mail.Read | |
| `mail_get_message` | Mail.Read | |
| `mail_list_folders` | Mail.Read | |
| `mail_list_attachments` | Mail.Read | |
| `mail_send_message` | Mail.Send | ✏️ |
| `mail_create_draft` | Mail.ReadWrite | ✏️ |
| `mail_reply_message` | Mail.Send | ✏️ |
| `mail_delete_message` | Mail.ReadWrite | ✏️ |
</details>

<details>
<summary><strong>📅 Calendar</strong> — 8 tools</summary>

| Tool | Scopes | Writes |
|---|---|:---:|
| `calendar_list_events` | Calendars.Read | |
| `calendar_get_event` | Calendars.Read | |
| `calendar_list_calendars` | Calendars.Read | |
| `calendar_find_meeting_times` | Calendars.Read.Shared | |
| `calendar_get_free_busy` | Calendars.Read.Shared | |
| `calendar_create_event` | Calendars.ReadWrite | ✏️ |
| `calendar_update_event` | Calendars.ReadWrite | ✏️ |
| `calendar_delete_event` | Calendars.ReadWrite | ✏️ |
</details>

<details>
<summary><strong>👥 Contacts & People</strong> — 6 tools</summary>

| Tool | Scopes | Writes |
|---|---|:---:|
| `contacts_list` | Contacts.Read | |
| `contacts_search` | Contacts.Read | |
| `contacts_people_search` | People.Read | |
| `contacts_create` | Contacts.ReadWrite | ✏️ |
| `contacts_update` | Contacts.ReadWrite | ✏️ |
| `contacts_delete` | Contacts.ReadWrite | ✏️ |
</details>

<details>
<summary><strong>📁 Files (OneDrive + SharePoint)</strong> — 10 tools</summary>

| Tool | Scopes | Writes |
|---|---|:---:|
| `files_list_children` | Files.Read.All, Sites.Read.All | |
| `files_get_item` | Files.Read.All | |
| `files_search` | Files.Read.All | |
| `files_download` | Files.Read.All | |
| `files_list_drives` | Files.Read.All | |
| `sites_search` | Sites.Read.All | |
| `files_upload` | Files.ReadWrite.All | ✏️ |
| `files_create_folder` | Files.ReadWrite.All | ✏️ |
| `files_delete` | Files.ReadWrite.All | ✏️ |
| `files_copy` | Files.ReadWrite.All | ✏️ |
| `files_share` | Files.ReadWrite.All | ✏️ |
</details>

<details>
<summary><strong>💬 Teams</strong> — 9 tools</summary>

| Tool | Scopes | Writes |
|---|---|:---:|
| `teams_list_joined` | Team.ReadBasic.All | |
| `teams_list_channels` | Channel.ReadBasic.All | |
| `teams_list_channel_messages` | ChannelMessage.Read.All | |
| `teams_get_message_replies` | ChannelMessage.Read.All | |
| `teams_list_chats` | Chat.Read | |
| `teams_list_chat_messages` | Chat.Read | |
| `teams_post_channel_message` | ChannelMessage.Send | ✏️ |
| `teams_reply_channel_message` | ChannelMessage.Send | ✏️ |
| `teams_post_chat_message` | ChatMessage.Send | ✏️ |
</details>

<details>
<summary><strong>✅ Tasks (To Do + Planner)</strong> — 8 tools</summary>

| Tool | Scopes | Writes |
|---|---|:---:|
| `todo_list_lists` | Tasks.Read | |
| `todo_list_tasks` | Tasks.Read | |
| `todo_create_task` | Tasks.ReadWrite | ✏️ |
| `todo_complete_task` | Tasks.ReadWrite | ✏️ |
| `planner_list_plans` | Tasks.Read | |
| `planner_list_tasks` | Tasks.Read | |
| `planner_create_task` | Tasks.ReadWrite | ✏️ |
| `planner_complete_task` | Tasks.ReadWrite | ✏️ |
</details>

<details>
<summary><strong>📓 OneNote</strong> — 6 tools</summary>

| Tool | Scopes | Writes |
|---|---|:---:|
| `onenote_list_notebooks` | Notes.Read | |
| `onenote_list_sections` | Notes.Read | |
| `onenote_list_pages` | Notes.Read | |
| `onenote_get_page_content` | Notes.Read | |
| `onenote_create_page` | Notes.ReadWrite | ✏️ |
| `onenote_delete_page` | Notes.ReadWrite | ✏️ |
</details>

<details>
<summary><strong>📊 Excel</strong> — 17 tools</summary>

| Tool | Scopes | Writes |
|---|---|:---:|
| `excel_create_session` | Files.ReadWrite.All | |
| `excel_close_session` | Files.ReadWrite.All | |
| `excel_list_worksheets` | Files.Read.All | |
| `excel_get_range` | Files.Read.All | |
| `excel_update_range` | Files.ReadWrite.All | ✏️ |
| `excel_list_tables` | Files.Read.All | |
| `excel_get_table_rows` | Files.Read.All | |
| `excel_add_table_rows` | Files.ReadWrite.All | ✏️ |
| `excel_run_workbook_calculation` | Files.ReadWrite.All | ✏️ |
| `excel_create_workbook` | Files.ReadWrite.All, Sites.ReadWrite.All | ✏️ |
| `excel_create_from_template` | Files.ReadWrite.All, Sites.ReadWrite.All | ✏️ |
| `excel_add_worksheet` | Files.ReadWrite.All | ✏️ |
| `excel_delete_worksheet` | Files.ReadWrite.All | ✏️ |
| `excel_rename_worksheet` | Files.ReadWrite.All | ✏️ |
| `excel_create_table` | Files.ReadWrite.All | ✏️ |
| `excel_set_formula` | Files.ReadWrite.All | ✏️ |
| `excel_clear_range` | Files.ReadWrite.All | ✏️ |
</details>

<details>
<summary><strong>📝 Word</strong> — 10 tools (create from scratch or from templates + OOXML in-place editing)</summary>

| Tool | Scopes | Writes |
|---|---|:---:|
| `word_read_text` | Files.Read.All | |
| `word_list_paragraphs` | Files.Read.All | |
| `word_replace_text` | Files.ReadWrite.All | ✏️ |
| `word_append_paragraph` | Files.ReadWrite.All | ✏️ |
| `word_create_document` | Files.ReadWrite.All, Sites.ReadWrite.All | ✏️ |
| `word_create_from_template` | Files.ReadWrite.All, Sites.ReadWrite.All | ✏️ |
| `word_append_heading` | Files.ReadWrite.All | ✏️ |
| `word_append_bullets` | Files.ReadWrite.All | ✏️ |
| `word_insert_paragraph_at` | Files.ReadWrite.All | ✏️ |
| `word_delete_paragraph` | Files.ReadWrite.All | ✏️ |
</details>

<details>
<summary><strong>🎞️ PowerPoint</strong> — 8 tools (create from scratch or from templates + OOXML in-place editing)</summary>

| Tool | Scopes | Writes |
|---|---|:---:|
| `powerpoint_list_slides` | Files.Read.All | |
| `powerpoint_get_slide_text` | Files.Read.All | |
| `powerpoint_extract_all_text` | Files.Read.All | |
| `powerpoint_replace_text` | Files.ReadWrite.All | ✏️ |
| `powerpoint_create_deck` | Files.ReadWrite.All, Sites.ReadWrite.All | ✏️ |
| `powerpoint_create_from_template` | Files.ReadWrite.All, Sites.ReadWrite.All | ✏️ |
| `powerpoint_add_slide` | Files.ReadWrite.All | ✏️ |
| `powerpoint_delete_slide` | Files.ReadWrite.All | ✏️ |
</details>

---

## Configuration

All settings can be passed as CLI flags **or** environment variables.

| Variable | Flag | Default | Notes |
|---|---|---|---|
| `MCP_AUTH_MODE` | `--auth` | `device-code` | `device-code` \| `client-credentials` \| `interactive` |
| `MCP_TENANT_ID` | `--tenant` | `common` | Tenant id, or `common` / `organizations` / `consumers` |
| `MCP_CLIENT_ID` | `--client-id` | Public Graph client | Use your own app for production |
| `MCP_CLIENT_SECRET` | `--client-secret` | – | Required for `client-credentials` |
| `MCP_REDIRECT_URI` | `--redirect-uri` | `http://localhost:3000` | `interactive` only |
| `MCP_SCOPES` | `--scopes` | Full set | Comma-separated Graph scopes |
| `MCP_ENABLE_WRITES` | `--enable-writes` | `false` | Unlock all mutating tools |
| `MCP_ENABLE_<SURFACE>_WRITE` | – | `false` | e.g. `MCP_ENABLE_MAIL_WRITE`, `MCP_ENABLE_FILES_WRITE` |
| `MCP_DISABLED_TOOLS` | `--disabled-tools` | – | CSV of tool names to hide entirely |
| `MCP_TOKEN_CACHE_PATH` | `--token-cache` | `~/.microsoft365-mcp/tokencache.json` | |
| `MCP_LOG_LEVEL` | – | `info` | Pino log level (stderr) |

---

## Known limitations

Most "this doesn't work" reports on business tenants aren't bugs in this server — they're tenant policy, auth-flow constraints, or current implementation scope. Read this before filing an issue.

### Tenant policy (your IT admin controls these — not a bug)

- **App registration blocked.** Many orgs disable end-user app registration. You'll see *"You do not have permission to create app registrations"* in Entra. → Ask IT to register one app and share the `MCP_CLIENT_ID` / `MCP_TENANT_ID`.
- **User consent blocked.** Some tenants require admin consent for *all* Graph scopes. Device-code sign-in will fail with `AADSTS65001` or `AADSTS90094`. → A Global Admin must click **"Grant admin consent for \<tenant\>"** on the app's API permissions page.
- **Device-code flow disabled.** Tenants can block device-code entirely via Conditional Access (`AADSTS50199` / `AADSTS530032`). → Use `MCP_AUTH_MODE=interactive` (local browser) or `client-credentials` (unattended) instead.
- **Conditional Access / MFA / compliant-device requirements.** CA policies can reject the sign-in even when everything else is configured correctly (unmanaged device, missing MFA, non-corporate network). The error comes from Entra, not this server. → Work with IT on an exemption or sign in from a compliant device.
- **"Allow public client flows" = No.** Required for device-code and interactive. If IT leaves this off, both user flows fail. → Enable it in **Authentication → Advanced settings**, or use `client-credentials`.
- **Guest / B2B accounts.** Guests in a tenant often have reduced Graph reach (e.g. can't read the directory). Expect some tools to return empty results even when they "should" work.
- **Personal Microsoft accounts + `client-credentials`.** Not supported by Microsoft — `client-credentials` is tenant-only. Personal accounts must use `device-code` or `interactive`.

### Scope & data access

- **Ungranted scopes fail silently per-tool.** If you didn't grant `Mail.Read`, `mail_list_messages` returns a Graph `Authorization_RequestDenied` error but the server stays up and other tools keep working. This is intentional — check the error payload, don't assume the server is broken.
- **Application permissions have tenant-wide reach by default.** `Mail.Read` (Application) reads *every* mailbox in the tenant. Always narrow with [Application Access Policies](https://learn.microsoft.com/graph/auth-limit-mailbox-access) (Mail), [RBAC for Applications](https://learn.microsoft.com/exchange/permissions-exo/application-rbac) (Exchange), or [`Sites.Selected`](https://learn.microsoft.com/sharepoint/dev/solution-guidance/security-apponly-azureacs) (SharePoint).
- **Shared / delegated mailboxes.** `Mail.Read.Shared` is required to read another user's mailbox you've been delegated; it's not granted by default.
- **Planner vs. Project.** Only Microsoft Planner is supported via Graph. Microsoft Project / Project for the Web is a different API surface and not covered.
- **Microsoft's built-in template gallery is not reachable.** The stock templates that appear in Office apps under *File → New* (*Ion Boardroom*, *Berlin*, *Modern proposal*, etc.) are served by `templates.office.com`, not Microsoft Graph — so `powerpoint_create_from_template` / `word_create_from_template` / `excel_create_from_template` cannot pull from that gallery. The one-time workaround: open the app → *File → New* → pick the template → save it to OneDrive, e.g. `/Templates/Ion-Boardroom.pptx`. After that it's addressable like any other file (`templatePath` or `templateItemId`, or resolved by `files_search`).

### Auth & token cache

- **Token cache is per-OS-user.** Cached at `~/.microsoft365-mcp/tokencache.json`. Switching Microsoft accounts means deleting that file first — there's no in-product account switcher.
- **Refresh tokens expire.** After ~90 days of inactivity (or on password change / revocation) the next call will prompt for re-auth via device code. Not a bug.
- **`keytar` is optional.** On headless Linux without a keyring, install fails gracefully and the cache falls back to the `chmod 600` JSON file. No encryption at rest in that case — acceptable for a personal machine, not for shared hosts.
- **First-run device-code sign-in is surfaced two ways, both in-client.** When there's no cached token, most MCP clients — including Claude Desktop — **do not render a server's stderr in the chat**, so if we only wrote the code there, the first tool call would appear to hang with no visible prompt. To work around that, the server:
  1. Kicks off the device-code flow at **startup** (not on first tool call) so the prompt is ready before you invoke anything.
  2. Emits the URL + code as an **MCP `notifications/message`** (logging channel) so clients that render server logs can show it in-UI.
  3. If a tool is called while sign-in is still pending, the tool returns an `isError` response containing the URL + code as its text — so even clients that *don't* render logging notifications will show the sign-in instructions inline in the chat. After signing in, re-run the tool and the cached token is used.

  All three paths co-exist; the stderr write is still there for CLI users and log-tailers (`%APPDATA%\Claude\logs\mcp-server-microsoft365.log` on Windows).

### Current implementation scope

These are real gaps in *this* server, tracked in [Roadmap](#roadmap):

- **Large file uploads (>4 MB) not yet supported.** `files_upload` uses the single-shot endpoint. Bigger files return a 413 from Graph. Upload sessions are on the roadmap.
- **Downloads are base64-encoded in the tool result.** Fine for small files; memory-heavy for large ones. Streaming/resource-URI downloads are planned.
- **No webhook / change-notification tools.** You can't subscribe to mailbox or drive changes — only poll.
- **No built-in retry on Graph 429/503.** Throttling errors surface directly to the caller. Retry-with-backoff is on the roadmap; for now, re-invoke the tool.
- **OOXML edits (Word / PowerPoint) download → mutate → re-upload.** No partial updates. Large decks mean large round-trips; concurrent edits by a human in the web app can be overwritten.
- **Excel requires a workbook session for writes.** You must call `excel_create_session` first — this is a Graph requirement, not something we can abstract away.

---

## Security

- **Read-only default.** Every mutating tool is hidden from `tools/list` until writes are enabled. Defense in depth: the handler re-checks at call time.
- **Token storage.** Cached locally with `chmod 600` (POSIX). On platforms with `keytar`, the cache is encrypted by the OS keychain. No tokens are logged — logger redacts common token fields.
- **Least-privilege scopes.** Override `MCP_SCOPES` to request only what you need. Tools whose scopes aren't granted simply won't succeed — the server stays up.
- **Bring-your-own app registration.** The default public `MCP_CLIENT_ID` is convenient for trying things out; **don't use it in production**. Register your own app and set `MCP_CLIENT_ID` / `MCP_TENANT_ID`.
- **Responsible disclosure:** see [SECURITY.md](./SECURITY.md).

---

## Compatibility

Tested with:

- ✅ Claude Desktop (macOS, Windows)


Any MCP-compliant client should work — open an issue if yours doesn't.

---

## Development

```bash
git clone https://github.com/lucasgfsvd/microsoft365-mcp-server
cd microsoft365-mcp-server
npm install
npm run dev   # builds + starts on stdio

# Verify tool shapes
npx @modelcontextprotocol/inspector node dist/index.js
```

Run tests: `npm test` · Typecheck: `npm run typecheck` · Lint: `npm run lint`.

See [CONTRIBUTING.md](./CONTRIBUTING.md) — contributions welcome, especially new surface coverage and integration test fixtures.

---

## Roadmap

- [ ] Large-file upload sessions (>4 MB)
- [ ] Streaming downloads (return resource URIs instead of base64)
- [ ] Subscription/webhook tools (real-time change notifications)
- [ ] MCP Resources for notebooks, sites, and mailboxes
- [ ] Prompt templates for common workflows (triage, meeting-prep)
- [ ] Loop/Whiteboard/Viva surfaces when Graph exposes them
- [ ] Per-tool rate-limit & retry with backoff surfaced in responses

---

## License

[MIT](./LICENSE)

## Acknowledgments

Built on [Microsoft Graph](https://learn.microsoft.com/graph/), [`@azure/identity`](https://github.com/Azure/azure-sdk-for-js), and the [Model Context Protocol SDK](https://github.com/modelcontextprotocol). Inspired by the growing MCP ecosystem and the many community servers that paved the way.
