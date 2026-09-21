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

1. Start the server (via npx, Docker, or `npm run dev`). It starts **signed out and silent** — starting it never triggers a prompt.
2. Call the **`auth_sign_in`** tool. It returns immediately with the URL and code:

   ```json
   {
     "status": "prompt-issued",
     "verificationUri": "https://login.microsoft.com/device",
     "userCode": "A1B2C3D4E"
   }
   ```

   The same code goes to stderr and to the MCP logging channel, for CLI users and clients that render server logs.
3. Visit the URL, enter the code, sign in, grant consent.
4. Call **`auth_status`** to confirm — it should report `"signedIn": true`.
5. A refreshable token is cached (OS keychain where available, otherwise a `chmod 600` file under `~/.microsoft365-mcp/`), along with an `authrecord.json` that records which account to reuse. Later starts authenticate silently.

> Any Graph tool invoked while signed out returns an error telling you to run `auth_sign_in`, rather than blocking. Neither auth tool is mutating, so both work with writes disabled.

### Bring your own app registration (recommended beyond "try it out")

The default `MCP_CLIENT_ID` is the public "Microsoft Graph Command Line Tools" app, which is fine for experimentation but noisy in audit logs. For real use:

1. Entra ID → App registrations → New registration.
2. Supported account types — pick the narrowest that fits:
   - *Accounts in this organizational directory only (Single tenant)* for a work/school account. Also set `MCP_TENANT_ID` to your **Directory (tenant) ID** rather than leaving it as `common`.
   - *…any organizational directory and personal Microsoft accounts* only if you need an `@outlook.com` account too. Keep `MCP_TENANT_ID=common`.
3. Leave **Redirect URI** blank. Device-code flow does not use one. (Only `MCP_AUTH_MODE=interactive` does — that needs a **Mobile and desktop applications** redirect of `http://localhost:3000`.)
4. Under **Authentication → Advanced settings**, set *Allow public client flows* = **Yes**. Device code is a public-client flow; without this, sign-in fails with `AADSTS7000218`.
5. API permissions → Add → **Delegated** — add the scopes listed in [permissions.md](./permissions.md) for the surfaces you want, then **Grant admin consent**. Some Teams scopes require it, and granting up front avoids per-user consent prompts.
6. Copy the **Application (client) ID** into `MCP_CLIENT_ID`.

Neither the client id nor the tenant id is a secret — a public client has no secret, and the tenant id is discoverable from your domain's OIDC metadata. Still, keep them out of public repositories: a client id is enough to build a consent-phishing link carrying your app's name.

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
