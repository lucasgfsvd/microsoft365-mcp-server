# Authentication

This server supports three auth flows. Pick the one that matches your situation.

| Flow | Audience | Requires | Interactivity |
|---|---|---|---|
| `device-code` (default) | Individuals, dev loops | Public client id | Once (device code in a browser) |
| `client-credentials` | Tenant-wide automation | Tenant id + app id + secret + admin consent | None |
| `interactive` | Desktop users preferring a browser | Public client id, reachable localhost redirect | Once (browser popup) |

---

## 1. Device code (individuals)

**Best for:** personal Microsoft accounts, single users, trying the server out.

1. Start the server (via npx, Docker, or `npm run dev`).
2. On the first tool call, the server prints to **stderr**:

   ```
   [microsoft365-mcp] To sign in, open https://microsoft.com/devicelogin and enter code A1B2C3D4E
   ```

3. Visit the URL, enter the code, sign in, grant consent.
4. A refreshable token is cached at `~/.microsoft365-mcp/tokencache.json` (`chmod 600`). Subsequent runs are silent.

### Bring your own app registration (recommended beyond "try it out")

The default `MCP_CLIENT_ID` is the public "Microsoft Graph Command Line Tools" app, which is fine for experimentation but noisy in audit logs. For real use:

1. Entra ID → App registrations → New registration.
2. Supported account types: *Accounts in any organizational directory and personal Microsoft accounts*.
3. Add platform: **Mobile and desktop applications**. Redirect URI: `http://localhost` (device-code flow uses this as a sentinel).
4. Under **Authentication**, enable *Allow public client flows* = Yes.
5. API permissions → Add → **Delegated** — add the scopes listed in [permissions.md](./permissions.md) for the surfaces you want.
6. Copy the **Application (client) ID** into `MCP_CLIENT_ID`.

No client secret is needed for device-code flow.

---

## 2. Client credentials (tenants)

**Best for:** companies running the server as unattended automation, shared across users, or hosted on a server.

1. Entra ID → App registrations → New registration. Single-tenant is typical.
2. Certificates & secrets → New client secret. Save the value (it's shown once).
3. API permissions → **Application** permissions (not Delegated) — add `.All` variants of what you need (e.g. `Mail.Read`, `Files.ReadWrite.All`). See [permissions.md](./permissions.md).
4. Click **Grant admin consent**. An admin must approve.
5. Run the server with:

   ```bash
   MCP_AUTH_MODE=client-credentials \
   MCP_TENANT_ID=<tenant-guid> \
   MCP_CLIENT_ID=<app-id> \
   MCP_CLIENT_SECRET=<secret> \
   npx -y @microsoft365-mcp/server
   ```

**Important:** application permissions grant tenant-wide access. Scope the app to specific mailboxes with [application access policies](https://learn.microsoft.com/graph/auth-limit-mailbox-access) if you don't want full-tenant reach.

---

## 3. Interactive browser

**Best for:** desktop developers who prefer a browser popup over typing a code.

1. Follow the same app registration steps as device-code, but set the redirect URI to a reachable `http://localhost:<port>` (e.g. `http://localhost:3000`) as a **Web** platform redirect.
2. Run with `MCP_AUTH_MODE=interactive` and optionally `MCP_REDIRECT_URI=http://localhost:3000`.

---

## Revoking access

- Individual device-code / interactive: go to [myapps.microsoft.com](https://myapps.microsoft.com) → **Manage your apps** → remove the app. Then delete `~/.microsoft365-mcp/tokencache.json`.
- Tenant client-credentials: rotate/delete the client secret in Entra ID, or remove API permissions.
