import { loadConfig } from "./config.js";
import { startServer } from "./server.js";
import { logger } from "./util/logger.js";

async function main(): Promise<void> {
  const config = loadConfig(process.argv);
  await startServer(config);
}

main().catch((err) => {
  logger.error({ err }, "fatal");
  process.stderr.write(`\n[microsoft365-mcp] fatal: ${(err as Error).message}\n`);
  process.exit(1);
});
