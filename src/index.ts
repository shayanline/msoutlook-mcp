#!/usr/bin/env node
/**
 * msoutlook-mcp — MCP server for Microsoft Outlook Web
 *
 * No app registration required. Uses your existing Outlook Web session,
 * the same way msteams-mcp uses the Microsoft Teams web session.
 *
 * Usage: npx msoutlook-mcp
 */

import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { createServer } from './server.js';
import { logger } from './utils/logger.js';

async function main(): Promise<void> {
  const server = createServer();
  const transport = new StdioServerTransport();

  await server.server.connect(transport);
  logger.info('msoutlook-mcp running on stdio');
}

main().catch(err => {
  logger.error('Fatal error', err);
  process.exit(1);
});
