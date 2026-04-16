# Security policy

## Supported versions

We support the latest `0.x` release on the `main` branch.

## Reporting a vulnerability

**Please do not file public GitHub issues for security reports.**

Instead:

- Use [GitHub private security advisories](https://github.com/lucasgfsvd/microsoft365-mcp-server/security/advisories/new), or
- Email the maintainers listed in `package.json` with a reproducible description.

We aim to acknowledge reports within 72 hours and to publish a fix and coordinated disclosure within 30 days for high-severity issues.

## Scope

In scope:

- Token handling, cache storage, and logging redaction
- Authentication flow correctness
- Injection / SSRF in tool handlers
- Supply chain (our dependencies and build)

Out of scope:

- Vulnerabilities in Microsoft Graph itself (report to Microsoft MSRC)
- Social engineering of users into enabling writes
- Missing rate-limiting — Microsoft Graph enforces its own limits

## Hardening guidance

- Always run with your **own** Azure AD app registration in production, not the default public client id.
- Keep writes disabled unless a specific tool requires them; prefer per-surface enablement.
- Mount the token cache on encrypted storage; on multi-user machines, ensure `chmod 600` is honored.
- Rotate client secrets regularly and use [Workload Identity Federation](https://learn.microsoft.com/entra/workload-id/workload-identity-federation) instead where possible.
