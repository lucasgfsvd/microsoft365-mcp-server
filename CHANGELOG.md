# Changelog

## [0.1.0] - unreleased

First and only planned release, published once the "Before 0.1.0" items in
`docs/roadmap.md` are done (add them below as they land, and date this heading
when tagging). After it, the project is provided as-is and not actively
maintained: see *Maintenance status* in the README.

### What it does

- **102 tools across 12 surfaces**: Mail, Calendar, Contacts and People,
  OneDrive and SharePoint files, Teams, To Do and Planner, OneNote, Excel,
  Word and PowerPoint, plus sign-in and cross-cutting Graph tools.
- **Writes are off by default.** Enable them all, or per surface; sends and
  creates are marked so clients can ask for confirmation.
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
- **Large files**: uploads over 4 MB use Graph upload sessions; large
  downloads stream to a folder you configure (`MCP_DOWNLOAD_DIR`) instead of
  entering the conversation.

### Safety

- A `POST` that sends or creates something is retried only on 429, so a
  transient 503/504 cannot send the same mail twice.
- Pre-authenticated download links (`tempauth` URLs) are stripped from every
  tool result.
- Sign-in is lazy and explicit (`auth_sign_in`); nothing prompts on startup.
  `--logout` removes this app's tokens from the token store and leaves other
  apps' alone.
- Runs where no OS keyring is available, falling back to an in-memory token
  cache.

### Tested

All 102 tools exercised against a real Microsoft 365 business tenant
(`scripts/live/`), with edited Office files also opened by independent
parsers; 206 unit tests. Teams channel posts and Planner tasks were tested in
a private team with no other members.

### Known limits

- `files_upload` takes files up to about 47 MB with the default
  `MCP_MAX_MESSAGE_MB` of 64, since content travels base64-encoded in the
  tool call.
- No change notifications (webhooks); use `graph_delta`. See `docs/webhooks.md`.
- In the Docker image, device-code sign-in does not survive a restart (no OS
  keyring); client-credentials mode is unaffected.
