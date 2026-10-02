import type { AccessToken, AuthenticationRecord, GetTokenOptions, TokenCredential } from "@azure/identity";
import { logger } from "../util/logger.js";
import type { ServerConfig } from "../types.js";
import { buildCredential } from "./index.js";
import { isAuthenticationRequired } from "./session.js";
import { readAuthRecord } from "./tokenCache.js";

type Build = (record?: AuthenticationRecord) => Promise<TokenCredential>;
type Authenticate = (scopes: string | string[]) => Promise<AuthenticationRecord | undefined>;

const sameRecord = (a?: AuthenticationRecord, b?: AuthenticationRecord) =>
  !!a && !!b && a.homeAccountId === b.homeAccountId && a.tenantId === b.tenantId && a.clientId === b.clientId && a.authority === b.authority;

/**
 * A credential that notices a sign-in completed by another server process.
 *
 * A credential learns which cached account to use from the authentication
 * record it is built with, and only then. So a server started before the user
 * signed in elsewhere kept failing with "sign-in required" until restarted,
 * though the tokens were already in the shared cache. When a token request
 * finds no account, this re-reads authrecord.json and, if another process has
 * written a record since, rebuilds the credential with it and asks again.
 * Anything else, and a record it has already tried, fails as before.
 */
export class ReloadingCredential implements TokenCredential {
  private reloading?: Promise<boolean>;

  constructor(
    private inner: TokenCredential,
    private record: AuthenticationRecord | undefined,
    private readonly build: Build,
    private readonly readRecord: () => Promise<AuthenticationRecord | undefined>,
    private readonly clientId: string,
  ) {}

  async getToken(scopes: string | string[], options?: GetTokenOptions): Promise<AccessToken | null> {
    try {
      return await this.inner.getToken(scopes, options);
    } catch (err) {
      if (!isAuthenticationRequired(err)) throw err;
      // One reload at a time; concurrent callers share its outcome.
      this.reloading ??= this.reload().finally(() => (this.reloading = undefined));
      if (!(await this.reloading)) throw err;
      return this.inner.getToken(scopes, options);
    }
  }

  /** Sign in interactively through the current credential, which keeps the account it gets. */
  async authenticate(scopes: string | string[]): Promise<AuthenticationRecord | undefined> {
    const authenticate = (this.inner as { authenticate?: Authenticate }).authenticate;
    if (typeof authenticate !== "function") throw new Error("This credential has no interactive sign-in.");
    const record = await authenticate.call(this.inner, scopes);
    if (record) this.record = record;
    return record;
  }

  private async reload(): Promise<boolean> {
    const fresh = await this.readRecord().catch(() => undefined);
    // Another app's record, or one already in use, would not help.
    if (!fresh || fresh.clientId !== this.clientId || sameRecord(fresh, this.record)) return false;
    this.inner = await this.build(fresh);
    this.record = fresh;
    logger.info("auth: picked up a sign-in completed by another process");
    return true;
  }
}

/**
 * The credential a server runs with. User sign-in modes get the reloading
 * wrapper; client credentials have no account to pick up.
 */
export async function buildServerCredential(config: ServerConfig): Promise<TokenCredential> {
  // Reusing the stored record lets a fresh process spend the cached token without
  // prompting. Absent or unreadable, we simply start out signed-out.
  const readRecord = () => readAuthRecord(config.tokenCachePath);
  const record = await readRecord().catch(() => undefined);
  if (config.authMode === "client-credentials") return buildCredential(config);
  return new ReloadingCredential(await buildCredential(config, record), record, (r) => buildCredential(config, r), readRecord, config.clientId);
}
