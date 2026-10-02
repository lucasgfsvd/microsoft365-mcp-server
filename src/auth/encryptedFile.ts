import { createCipheriv, createDecipheriv, randomBytes, scrypt } from "node:crypto";
import { promises as fs } from "node:fs";
import path from "node:path";

/**
 * A small file of secret text, encrypted with a key derived from a passphrase.
 *
 * AES-256-GCM, so a wrong passphrase or a tampered file is detected rather than
 * read as garbage. The key comes from scrypt over a random per-file salt, which
 * makes guessing a weak passphrase costly. Writes go to a temporary file and are
 * renamed into place, so a crash cannot leave half a file.
 */

const FORMAT = "microsoft365-mcp/encrypted-token-cache";
const VERSION = 1;
/** Binds the ciphertext to its purpose, so it cannot be passed off as something else. */
const AAD = Buffer.from(`${FORMAT}/v${VERSION}`);
const SCRYPT = { N: 2 ** 15, r: 8, p: 1, maxmem: 64 * 1024 * 1024 };

interface Envelope {
  format: typeof FORMAT;
  v: typeof VERSION;
  salt: string;
  iv: string;
  tag: string;
  data: string;
}

/** The file exists but cannot be read with this passphrase (or is not ours). */
export class UnreadableFileError extends Error {
  constructor(file: string, why: string) {
    super(`Cannot read ${file}: ${why}`);
    this.name = "UnreadableFileError";
  }
}

export class EncryptedFile {
  private readonly keys = new Map<string, Promise<Buffer>>();
  /** Salt of the file as last read or written; kept so each write does not re-derive the key. */
  private salt?: Buffer;

  constructor(
    readonly file: string,
    private readonly passphrase: string,
  ) {}

  private key(salt: Buffer): Promise<Buffer> {
    const id = salt.toString("base64");
    let key = this.keys.get(id);
    if (!key) {
      key = new Promise((resolve, reject) =>
        scrypt(this.passphrase, salt, 32, SCRYPT, (err, k) => (err ? reject(err) : resolve(k))),
      );
      this.keys.set(id, key);
    }
    return key;
  }

  /** The plaintext, or undefined when there is no file. Throws UnreadableFileError otherwise. */
  async read(): Promise<string | undefined> {
    let raw: string;
    try {
      raw = await fs.readFile(this.file, "utf8");
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code === "ENOENT") return undefined;
      throw err;
    }
    let env: Envelope;
    try {
      env = JSON.parse(raw) as Envelope;
    } catch {
      throw new UnreadableFileError(this.file, "not an encrypted token cache");
    }
    if (env.format !== FORMAT || env.v !== VERSION) throw new UnreadableFileError(this.file, "not an encrypted token cache");
    const salt = Buffer.from(env.salt, "base64");
    try {
      const decipher = createDecipheriv("aes-256-gcm", await this.key(salt), Buffer.from(env.iv, "base64"));
      decipher.setAAD(AAD);
      decipher.setAuthTag(Buffer.from(env.tag, "base64"));
      const text = Buffer.concat([decipher.update(Buffer.from(env.data, "base64")), decipher.final()]).toString("utf8");
      this.salt = salt;
      return text;
    } catch {
      throw new UnreadableFileError(this.file, "MCP_TOKEN_CACHE_KEY does not match the key it was written with");
    }
  }

  async write(text: string): Promise<void> {
    const salt = (this.salt ??= randomBytes(16));
    const iv = randomBytes(12);
    const cipher = createCipheriv("aes-256-gcm", await this.key(salt), iv);
    cipher.setAAD(AAD);
    const data = Buffer.concat([cipher.update(text, "utf8"), cipher.final()]);
    const env: Envelope = {
      format: FORMAT,
      v: VERSION,
      salt: salt.toString("base64"),
      iv: iv.toString("base64"),
      tag: cipher.getAuthTag().toString("base64"),
      data: data.toString("base64"),
    };
    await fs.mkdir(path.dirname(this.file), { recursive: true });
    const tmp = `${this.file}.${process.pid}.${randomBytes(4).toString("hex")}.tmp`;
    try {
      await fs.writeFile(tmp, JSON.stringify(env), { mode: 0o600 });
      await fs.rename(tmp, this.file);
    } catch (err) {
      await fs.unlink(tmp).catch(() => undefined);
      throw err;
    }
  }

  async remove(): Promise<boolean> {
    try {
      await fs.unlink(this.file);
      return true;
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code === "ENOENT") return false;
      throw err;
    }
  }
}
