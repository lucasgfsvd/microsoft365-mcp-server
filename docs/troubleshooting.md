# Troubleshooting

## The client says the server disconnected immediately

- Make sure Node.js ≥ 20 is on your PATH if you use the `npx` install.
- On Windows, Claude Desktop launches from a `cmd` shell — check the path it expects in logs.
- Run `npx -y @microsoft365-mcp/server --help` manually. If that hangs waiting for input, it's working (stdio mode).

## "AADSTS65001: The user or administrator has not consented…"

Your tenant requires admin consent for the scopes you asked for. Either:

- Switch to your own app registration and have an admin click **Grant admin consent**, or
- Narrow `MCP_SCOPES` to a subset the user can self-consent to.

## "Insufficient privileges to complete the operation"

The token is valid but the scope you need for that specific tool is missing. Check [permissions.md](./permissions.md), add the scope in the app registration, and re-consent (delete the cache to force a new login).

## Device code keeps prompting me

- Delete `~/.microsoft365-mcp/tokencache.json` and try again.
- Verify the cache file is writable (`chmod 600`, owned by your user).
- Check system time is correct — token validation is time-sensitive.

## Writes don't work

By design, writes are off by default. Set `MCP_ENABLE_WRITES=true` (or a per-surface flag like `MCP_ENABLE_MAIL_WRITE=true`) and restart the server.

## "Tool not found" in the MCP client

The MCP client caches tool lists at startup. After enabling writes, **restart your client** so it re-queries `tools/list`.

## Rate limiting / 429 responses

Microsoft Graph throttles per-tenant and per-app. Back off, reduce `top` page sizes, and avoid polling loops. For heavy Excel editing, always use `excel_create_session` to avoid per-call reload overhead.

## Large files

`files_upload` is capped at 4 MB in v0.1. Upload sessions for larger files are on the roadmap — track the issue or contribute a PR.
