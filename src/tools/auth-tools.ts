/**
 * Authentication MCP tools.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { browserLogin, getAuthStatus, clearSession, getOwaToken } from '../auth/index.js';

export function registerAuthTools(server: McpServer): void {
  // ── outlook_login ────────────────────────────────────────────────────────
  server.tool(
    'outlook_login',
    'Sign in to Outlook Web. Opens a browser only if not already authenticated — once signed in, the session is saved and reused automatically.',
    {},
    async () => {
      // Check if we already have valid (or refreshable) tokens — skip the browser if so
      const existingToken = await getOwaToken();
      if (existingToken) {
        const status = getAuthStatus();
        return {
          content: [{
            type: 'text',
            text: `Already signed in as ${status.upn ?? 'unknown'}. Token valid for ~${status.owaTokenMinutesRemaining ?? 0} more minutes. No browser needed.`,
          }],
        };
      }

      // No valid token — open the browser for fresh login.
      // Tell the user upfront: do NOT close the browser window, it closes itself.
      process.stderr.write('[msoutlook-mcp] Opening browser. Do NOT close the window — it will close automatically once signed in.\n');
      const upn = await browserLogin();
      return {
        content: [{
          type: 'text',
          text: upn
            ? `Signed in as ${upn}. Session saved — future calls will be silent (no browser).`
            : 'Login failed. If you closed the browser window manually, please run outlook_login again and wait for the browser to close on its own.',
        }],
      };
    },
  );

  // ── outlook_status ───────────────────────────────────────────────────────
  server.tool(
    'outlook_status',
    'Check the current authentication status and token validity.',
    {},
    async () => {
      // Try to get a valid token (triggers refresh if needed)
      const token = await getOwaToken();

      if (!token) {
        return {
          content: [{ type: 'text', text: 'Not authenticated. Run outlook_login to sign in.' }],
        };
      }

      const status = getAuthStatus();

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
