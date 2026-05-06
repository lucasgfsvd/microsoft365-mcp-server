import pino from "pino";

const REDACT_KEYS = [
  "access_token",
  "accessToken",
  "refresh_token",
  "refreshToken",
  "id_token",
  "idToken",
  "client_secret",
  "clientSecret",
  "client_assertion",
  "clientAssertion",
  "authorization",
  "Authorization",
];

export const logger = pino({
  level: process.env.MCP_LOG_LEVEL ?? "info",
  // MCP uses stdout for the protocol; logs must go to stderr.
  base: { service: "microsoft365-mcp-server" },
  redact: {
    paths: REDACT_KEYS.flatMap((k) => [k, `*.${k}`, `*.*.${k}`]),
    censor: "[redacted]",
  },
  transport: undefined,
}, pino.destination(2));
