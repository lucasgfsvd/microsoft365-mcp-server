import type { AuthenticationRecord, TokenCredential } from "@azure/identity";
import { logger } from "../util/logger.js";
import type { ServerConfig } from "../types.js";
import {
  clearLastDeviceCodePrompt,
  deviceCodeEmitter,
  getLastDeviceCodePrompt,
  type DeviceCodePrompt,
} from "./index.js";
import { writeAuthRecord } from "./tokenCache.js";

const GRAPH_DEFAULT_SCOPE = "https://graph.microsoft.com/.default";
/** How long signIn() waits for a device code before handing control back. */
const PROMPT_WAIT_MS = 5_000;

export type AuthState = "unknown" | "signed-in" | "sign-in-required";

export interface AuthStatus {
  state: AuthState;
  authMode: string;
  signedIn: boolean;
  pendingPrompt?: DeviceCodePrompt;
}

export interface SignInResult {
  status: "already-signed-in" | "prompt-issued" | "completed" | "failed";
  message: string;
  verificationUri?: string;
  userCode?: string;
}

/** Credentials that support an explicit interactive sign-in step. */
interface InteractiveCredential {
  authenticate?: (scopes: string | string[]) => Promise<AuthenticationRecord | undefined>;
}

/** True for the error Azure Identity raises when user interaction is required. */
export function isAuthenticationRequired(err: unknown): boolean {
  return (err as { name?: string } | undefined)?.name === "AuthenticationRequiredError";
}

/** Resolve when a device code arrives, or when the timeout elapses. */
function waitForPrompt(ms: number): Promise<void> {
  return new Promise((resolve) => {
    const done = (): void => {
      clearTimeout(timer);
      deviceCodeEmitter.off("prompt", done);
      resolve();
    };
    const timer = setTimeout(done, ms);
    deviceCodeEmitter.once("prompt", done);
  });
}

/**
 * Owns the credential's sign-in lifecycle.
 *
 * Authentication is lazy on purpose. Nothing here ever starts an interactive
 * flow on its own: `probe()` is silent and only ever spends a cached token, and
 * the device-code flow runs solely when `signIn()` is called. That is what keeps
 * a sign-in prompt from firing on every MCP client launch — the previous eager
 * warmup issued a fresh device code at startup, which then expired unused.
 */
export class AuthSession {
  private state: AuthState = "unknown";
  private interactive?: Promise<void>;

  constructor(
    private readonly credential: TokenCredential,
    private readonly config: ServerConfig,
  ) {}

  private get scopes(): string | string[] {
    return this.config.scopes.length > 0 ? this.config.scopes : GRAPH_DEFAULT_SCOPE;
  }

  /**
   * Resolve a token without any user interaction. Returns false when sign-in is
   * required. Genuine failures (network, bad tenant) are rethrown so they are
   * not misreported to the user as "please sign in".
   */
  async probe(): Promise<boolean> {
    // Fast path for the per-call gate: a token we already spent successfully
    // is almost certainly still good, and this runs before every Graph tool.
    if (this.state === "signed-in") return true;
    return this.verify();
  }

  /**
   * Ask the credential outright, ignoring what we last believed.
   *
   * `signed-in` is only ever an assumption: it records that a token was
   * obtained once, not that one can be obtained now. Trusting it in status()
   * meant auth_status could keep reporting `signedIn: true` while every Graph
   * call failed — which sends you looking for a bug in the wrong place.
   */
  private async verify(): Promise<boolean> {
    try {
      const token = await this.credential.getToken(this.scopes);
      if (!token) {
        this.state = "sign-in-required";
        return false;
      }
      this.markSignedIn();
      return true;
    } catch (err) {
      if (!isAuthenticationRequired(err)) throw err;
      this.state = "sign-in-required";
      return false;
    }
  }

  async status(): Promise<AuthStatus> {
    let signedIn = false;
    try {
      signedIn = await this.verify();
    } catch (err) {
      logger.warn({ err }, "auth check failed");
    }
    return {
      state: this.state,
      authMode: this.config.authMode,
      signedIn,
      pendingPrompt: getLastDeviceCodePrompt(),
    };
  }

  /**
   * Start (or join) an interactive sign-in. Returns as soon as there is a device
   * code to show, rather than blocking for the minutes a user needs to finish in
   * a browser. The flow continues in the background; `status()` reports the result.
   */
  async signIn(): Promise<SignInResult> {
    if (await this.probe()) {
      return { status: "already-signed-in", message: "Already signed in to Microsoft 365." };
    }

    const cred = this.credential as TokenCredential & InteractiveCredential;
    if (typeof cred.authenticate !== "function") {
      return {
        status: "failed",
        message: `Auth mode "${this.config.authMode}" authenticates without user interaction; nothing to sign in to.`,
      };
    }

    let pending = this.interactive;
    if (!pending) {
      clearLastDeviceCodePrompt();
      pending = cred
        .authenticate(this.scopes)
        .then(
          async (record) => {
            this.markSignedIn();
            if (record) await this.persistRecord(record);
            logger.info("interactive sign-in completed");
          },
          (err) => {
            this.state = "sign-in-required";
            logger.warn({ err }, "interactive sign-in did not complete");
          },
        )
        .finally(() => {
          this.interactive = undefined;
        });
      this.interactive = pending;
    }

    await Promise.race([pending, waitForPrompt(PROMPT_WAIT_MS)]);

    if (this.state === "signed-in") {
      return { status: "completed", message: "Signed in to Microsoft 365." };
    }

    const prompt = getLastDeviceCodePrompt();
    if (prompt) {
      return {
        status: "prompt-issued",
        message:
          `Open ${prompt.verificationUri} and enter code ${prompt.userCode}. ` +
          `Then call auth_status to confirm — the token is cached afterwards.`,
        verificationUri: prompt.verificationUri,
        userCode: prompt.userCode,
      };
    }

    return {
      status: "failed",
      message: "Sign-in could not be started. Check the server log on stderr for details.",
    };
  }

  /**
   * Becoming signed in retires any device code we were showing. Without this the
   * code outlives the sign-in it belongs to, and status() keeps reporting a
   * pendingPrompt that nobody is waiting on.
   */
  private markSignedIn(): void {
    this.state = "signed-in";
    clearLastDeviceCodePrompt();
  }

  /** Persist the record so the next process start can authenticate silently. */
  private async persistRecord(record: AuthenticationRecord): Promise<void> {
    try {
      await writeAuthRecord(this.config.tokenCachePath, record);
    } catch (err) {
      logger.warn({ err }, "could not persist authentication record; next start will prompt again");
    }
  }
}
