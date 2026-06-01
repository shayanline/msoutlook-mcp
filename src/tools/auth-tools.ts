/**
 * Authentication MCP tools.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { browserLogin, isAuthenticated, getAuthStatus, clearSession } from '../auth/index.js';
import { getOwaToken } from '../auth/index.js';

export function registerAuthTools(server: McpServer): void {
  // ── outlook_login ────────────────────────────────────────────────────────
  server.tool(
    'outlook_login',
    'Open a browser window to log in to Outlook Web. Required before using any other tools.',
    {},
    async () => {
      const success = await browserLogin();
      return {
        content: [
          {
            type: 'text',
            text: success
              ? 'Successfully logged in to Outlook Web. Your session has been saved.'
              : 'Login failed or timed out. Please try again.',
          },
        ],
      };
    },
  );

  // ── outlook_status ───────────────────────────────────────────────────────
  server.tool(
    'outlook_status',
    'Check the current authentication status and token validity.',
    {},
    async () => {
      if (!isAuthenticated()) {
        return {
          content: [
            {
              type: 'text',
              text: 'Not authenticated. Run outlook_login to sign in.',
            },
          ],
        };
      }

      const status = getAuthStatus();

      // Try to get a valid token (triggers refresh if needed)
      const token = await getOwaToken();

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                authenticated: true,
                tokenValid: !!token,
                upn: status.upn,
                tenantId: status.tenantId,
                owaTokenExpiry: status.owaTokenExpiry,
                owaTokenMinutesRemaining: status.owaTokenMinutesRemaining,
                graphTokenExpiry: status.graphTokenExpiry,
              },
              null,
              2,
            ),
          },
        ],
      };
    },
  );

  // ── outlook_logout ───────────────────────────────────────────────────────
  server.tool(
    'outlook_logout',
    'Clear the saved Outlook session and tokens. You will need to run outlook_login again.',
    {},
    async () => {
      clearSession();
      return {
        content: [{ type: 'text', text: 'Session cleared. Run outlook_login to sign in again.' }],
      };
    },
  );
}
