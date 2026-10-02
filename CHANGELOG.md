# Changelog

## [0.1.0] - unreleased

First and only planned release; everything `docs/roadmap.md` planned for it
is below (date this heading when tagging). After it, the project is provided as-is and not actively
maintained: see *Maintenance status* in the README.

### What it does

- **102 tools across 12 surfaces**: Mail, Calendar, Contacts and People,
  OneDrive and SharePoint files, Teams, To Do and Planner, OneNote, Excel,
  Word and PowerPoint, plus sign-in and cross-cutting Graph tools.
- **Writes are off by default.** Enable them all, or per surface. Every tool
  carries the standard MCP annotations (read-only, destructive, idempotent,
  reaches other people), so a client can decide what to confirm.
- **Word, PowerPoint and Excel editing in place**: create documents, decks and
  workbooks, or edit existing ones (replace text, add and delete paragraphs
  and slides, ranges, tables, formulas).
- **Cross-surface Graph tools**: `graph_search` (one relevance-ranked query
  across mail, files, Teams), `graph_batch_get` (up to 20 reads in one round
  trip), and `graph_delta` (what changed since last time, deletions included).
- **Prompt templates**: `daily-brief`, `inbox-triage` and `meeting-prep`.
- **MCP resources**: attach mail messages, files and OneNote pages as context;
  Word and PowerPoint come back as text, Excel as CSV per sheet.
- **Mail**: list by date (`receivedAfter` / `receivedBefore`), and draft a
  reply inside its conversation without sending it
  (`mail_create_reply_draft`), which `inbox-triage` now uses.
- **Tasks and OneNote**: delete To Do and Planner tasks; create OneNote
  notebooks and sections, so an account with no notebook can start one.
- **Large files**: uploads over 4 MB use Graph upload sessions, and a file in
  a folder you configure (`MCP_UPLOAD_DIR`) uploads straight from disk at any
  size; large downloads stream to another (`MCP_DOWNLOAD_DIR`). Either way the
  bytes stay out of the conversation, the client sees progress, and the
  transfer can be cancelled.

### Safety

- A `POST` that sends or creates something is retried only on 429, so a
  transient 503/504 cannot send the same mail twice.
- Pre-authenticated download links (`tempauth` URLs) are stripped from every
  tool result.
- Sign-in is lazy and explicit (`auth_sign_in`); nothing prompts on startup.
  `--logout` removes this app's tokens from the token store and leaves other
  apps' alone.
- Runs where no OS keyring is available. With `MCP_TOKEN_CACHE_KEY` set, the
  token cache is a file encrypted with that key, so a device-code sign-in
  survives restarts there too (the Docker image included); without it, tokens
  are held in memory.
- A server already running picks up a sign-in completed by another process on
  its next call, without a restart.

### Tested

All 102 tools exercised against a real Microsoft 365 business tenant
(`scripts/live/`), with edited Office files also opened by independent
parsers; 243 unit tests. Teams channel posts and Planner tasks were tested in
a private team with no other members. Sign-in persistence and pickup were
tested with a real device-code sign-in, locally and in the Docker image.

### Known limits

- `files_upload` with `contentBase64` takes files up to about 47 MB with the
  default `MCP_MAX_MESSAGE_MB` of 64; use `localPath` for anything larger.
- No change notifications (webhooks); use `graph_delta`. See `docs/webhooks.md`.
