# Security policy

## Supported versions and response

**This project is not actively maintained.** There is no supported version in the
sense of one that receives fixes: `0.1.0` is published as-is. Reports are still
welcome and will be read when possible, but there is **no commitment** to
acknowledge them, fix them, or publish a release. If you depend on this server,
fork it so you can patch it yourself.

## Reporting a vulnerability

**Please do not file public GitHub issues for security reports.** Use
[GitHub private security advisories](https://github.com/lucasgfsvd/microsoft365-mcp-server/security/advisories/new)
with a reproducible description.

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
