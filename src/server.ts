import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  type Tool as McpTool,
} from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import type { ServerConfig } from "./types.js";
import { buildCredential, deviceCodeEmitter, type DeviceCodePrompt } from "./auth/index.js";
import { AuthSession, isAuthenticationRequired } from "./auth/session.js";
import { readAuthRecord } from "./auth/tokenCache.js";
import { buildGraphClient } from "./graph/client.js";
import { ToolRegistry } from "./tools/registry.js";
import { allTools } from "./tools/index.js";
import { assertAllowed } from "./util/writeGuard.js";
import { normalizeGraphError } from "./graph/errors.js";
import { logger } from "./util/logger.js";
import { DEFAULT_PUBLIC_CLIENT_ID } from "./config.js";
import { serializeResult } from "./util/redact.js";
import { registerPrompts } from "./prompts/index.js";

const SIGN_IN_HINT =
  "Not signed in to Microsoft 365. Call auth_sign_in to get a device code, enter it " +
  "in a browser, then retry this tool.";

export async function startServer(config: ServerConfig): Promise<void> {
  if (config.clientId === DEFAULT_PUBLIC_CLIENT_ID) {
    logger.warn(
      "Using the default public Microsoft client id. Convenient for trying things out, " +
        "but for production register your own Entra app and set MCP_CLIENT_ID — your own " +
        "audit trail and scope set won't be muddled with everyone else's.",
    );
  }
  logger.info(
    {
      authMode: config.authMode,
      tenantId: config.tenantId,
      enableWrites: config.enableWrites,
      scopes: config.scopes,
    },
    "auth configuration",
  );

  // Reusing the stored record lets a fresh process spend the cached token without
  // prompting. Absent or unreadable, we simply start out signed-out.
  const authenticationRecord = await readAuthRecord(config.tokenCachePath).catch(() => undefined);
  const credential = await buildCredential(config, authenticationRecord);
  const graph = buildGraphClient(credential, config.scopes);
  const auth = new AuthSession(credential, config);

  const registry = new ToolRegistry();
  registry.registerAll(allTools());

  const server = new Server(
    { name: "microsoft365-mcp-server", version: "0.1.0" },
    { capabilities: { tools: {}, prompts: {}, logging: {} } },
  );

  // Surface the device-code prompt as an MCP logging notification so clients
  // that render server logs can show it in-UI, not only in stderr.
  deviceCodeEmitter.on("prompt", (p: DeviceCodePrompt) => {
    server
      .sendLoggingMessage({
        level: "info",
        data: `Sign-in required: open ${p.verificationUri} and enter code ${p.userCode}`,
      })
      .catch(() => {
        /* clients that haven't subscribed will reject — ignore */
      });
  });

  registerPrompts(server, () => new Set(registry.list(config).map((t) => t.name)));

  server.setRequestHandler(ListToolsRequestSchema, async () => {
    const visible = registry.list(config);
    const tools: McpTool[] = visible.map((t) => ({
      name: t.name,
      description: t.description + (t.mutating ? " [MUTATING]" : ""),
      inputSchema: z.toJSONSchema(t.inputSchema, { target: "draft-7" }) as McpTool["inputSchema"],
    }));
    return { tools };
  });

  server.setRequestHandler(CallToolRequestSchema, async (req) => {
    const tool = registry.get(req.params.name);
    if (!tool) {
      return { isError: true, content: [{ type: "text", text: `Unknown tool: ${req.params.name}` }] };
    }

    try {
      assertAllowed(tool, config);
      const args = tool.inputSchema.parse(req.params.arguments ?? {});

      // The auth surface drives sign-in itself, so it must run un-gated.
      if (tool.surface !== "auth" && !(await auth.probe())) {
        return { isError: true, content: [{ type: "text", text: SIGN_IN_HINT }] };
      }

      const result = await tool.handler(args, { graph, credential, config, auth });
      return {
        content: [{ type: "text", text: serializeResult(result) }],
        ...(tool.mutating ? { _meta: { requires_confirmation: true } } : {}),
      };
    } catch (err) {
      if (isAuthenticationRequired(err)) {
        return { isError: true, content: [{ type: "text", text: SIGN_IN_HINT }] };
      }
      const norm = normalizeGraphError(err);
      logger.error({ err: norm, tool: tool.name }, "tool call failed");
      return {
        isError: true,
        content: [{ type: "text", text: `${norm.code}: ${norm.message}` }],
      };
    }
  });

  // An over-limit message makes the SDK close the transport, and with it the
  // process. Say so, rather than vanishing with exit code 0.
  server.onerror = (err) => logger.error({ err: err.message }, "MCP transport error");
  server.onclose = () => logger.warn("MCP transport closed; server exiting");

  const transport = new StdioServerTransport(process.stdin, process.stdout, {
    maxBufferSize: config.maxMessageBytes,
  });
  await server.connect(transport);
  logger.info({ toolCount: registry.list(config).length }, "microsoft365-mcp-server ready (stdio)");

  // Silent and non-blocking: this can only ever spend an already-cached token,
  // because the credential is built with disableAutomaticAuthentication. It will
  // never issue a device code, so starting the client stays quiet.
  void auth
    .probe()
    .then((ok) =>
      logger.info(
        { signedIn: ok },
        ok ? "using cached Microsoft 365 credentials" : "signed out; call auth_sign_in when needed",
      ),
    )
    .catch((err) => logger.warn({ err }, "auth probe failed"));
}
