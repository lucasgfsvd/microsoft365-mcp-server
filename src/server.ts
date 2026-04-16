import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  type Tool as McpTool,
} from "@modelcontextprotocol/sdk/types.js";
import { zodToJsonSchema } from "zod-to-json-schema";
import type { ServerConfig } from "./types.js";
import {
  buildCredential,
  deviceCodeEmitter,
  getLastDeviceCodePrompt,
  type DeviceCodePrompt,
} from "./auth/index.js";
import { buildGraphClient } from "./graph/client.js";
import { ToolRegistry } from "./tools/registry.js";
import { allTools } from "./tools/index.js";
import { assertAllowed } from "./util/writeGuard.js";
import { normalizeGraphError } from "./graph/errors.js";
import { logger } from "./util/logger.js";

const SIGN_IN_WAIT_MS = 1500;
const GRAPH_DEFAULT_SCOPE = "https://graph.microsoft.com/.default";

function signInErrorText(p: DeviceCodePrompt): string {
  return (
    `Sign-in required.\n\n` +
    `Open ${p.verificationUri} and enter code ${p.userCode}.\n` +
    `After signing in, re-run this tool — the token will be cached.`
  );
}

export async function startServer(config: ServerConfig): Promise<void> {
  const credential = buildCredential(config);
  const graph = buildGraphClient(credential, config.scopes);

  const registry = new ToolRegistry();
  registry.registerAll(allTools());

  const server = new Server(
    { name: "microsoft365-mcp-server", version: "0.1.0" },
    { capabilities: { tools: {}, logging: {} } },
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

  // Warm the credential up in the background for device-code mode so the prompt
  // fires at startup, not on the first tool call. If the token is already cached,
  // this resolves immediately. If not, userPromptCallback populates the prompt.
  let warmupComplete = config.authMode !== "device-code";
  const warmupPromise: Promise<void> =
    config.authMode === "device-code"
      ? credential
          .getToken(config.scopes.length > 0 ? config.scopes : GRAPH_DEFAULT_SCOPE)
          .then(
            () => {
              warmupComplete = true;
            },
            (err) => {
              warmupComplete = true;
              logger.warn({ err }, "auth warmup failed");
            },
          )
      : Promise.resolve();

  server.setRequestHandler(ListToolsRequestSchema, async () => {
    const visible = registry.list(config);
    const tools: McpTool[] = visible.map((t) => ({
      name: t.name,
      description: t.description + (t.mutating ? " [MUTATING]" : ""),
      inputSchema: zodToJsonSchema(t.inputSchema, { target: "jsonSchema7" }) as McpTool["inputSchema"],
    }));
    return { tools };
  });

  server.setRequestHandler(CallToolRequestSchema, async (req) => {
    const tool = registry.get(req.params.name);
    if (!tool) {
      return { isError: true, content: [{ type: "text", text: `Unknown tool: ${req.params.name}` }] };
    }

    // If device-code warmup is still pending, wait briefly for either a token
    // or a user-visible prompt, then surface the prompt as a tool error rather
    // than blocking the call for minutes with no visible progress.
    if (!warmupComplete) {
      await Promise.race([
        warmupPromise,
        new Promise<void>((resolve) => setTimeout(resolve, SIGN_IN_WAIT_MS)),
      ]);
      if (!warmupComplete) {
        const prompt = getLastDeviceCodePrompt();
        if (prompt) {
          return { isError: true, content: [{ type: "text", text: signInErrorText(prompt) }] };
        }
      }
    }

    try {
      assertAllowed(tool, config);
      const args = tool.inputSchema.parse(req.params.arguments ?? {});
      const result = await tool.handler(args, { graph, credential, config });
      return {
        content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
        ...(tool.mutating ? { _meta: { requires_confirmation: true } } : {}),
      };
    } catch (err) {
      const norm = normalizeGraphError(err);
      logger.error({ err: norm, tool: tool.name }, "tool call failed");
      return {
        isError: true,
        content: [{ type: "text", text: `${norm.code}: ${norm.message}` }],
      };
    }
  });

  const transport = new StdioServerTransport();
  await server.connect(transport);
  logger.info({ toolCount: registry.list(config).length }, "microsoft365-mcp-server ready (stdio)");
}
