/**
 * MCP server setup — registers all tools and starts the server.
 */

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerAuthTools } from './tools/auth-tools.js';
import { registerMailTools } from './tools/mail-tools.js';
import { registerCalendarTools } from './tools/calendar-tools.js';
import { registerContactTools } from './tools/contact-tools.js';

export function createServer(): McpServer {
  const server = new McpServer({
    name: 'msoutlook-mcp',
    version: '0.1.0',
  });

  registerAuthTools(server);
  registerMailTools(server);
  registerCalendarTools(server);
  registerContactTools(server);

  return server;
}
