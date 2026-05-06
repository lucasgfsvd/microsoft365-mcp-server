import { loadConfig } from "./config.js";
import { runListTools, runLogout } from "./cli.js";
import { startServer } from "./server.js";
import { logger } from "./util/logger.js";

async function main(): Promise<void> {
  const config = loadConfig(process.argv);
  if (config.listTools) {
    runListTools(config);
    return;
  }
  if (config.logout) {
    await runLogout(config);
    return;
  }
  await startServer(config);
}

main().catch((err) => {
  logger.error({ err }, "fatal");
  process.stderr.write(`\n[microsoft365-mcp] fatal: ${(err as Error).message}\n`);
  process.exit(1);
});
