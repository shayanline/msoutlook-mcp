/**
 * Authentication MCP tools.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { browserLogin, getAuthStatus, clearSession, getOwaToken } from '../auth/index.js';

/** Minimum minutes remaining before we consider the token "still valid" (skip re-login). */
const TOKEN_VALID_THRESHOLD_MINUTES = 10;

export function registerAuthTools(server: McpServer): void {
  // ── outlook_login ────────────────────────────────────────────────────────
  server.tool(
    'outlook_login',
    'Sign in to Outlook Web. Tries silently first (no browser); opens a browser only when the session has expired. Set force_new: true to force a full re-login.',
    {
      force_new: z.boolean().optional().describe(
        'Force a full re-login even if a session exists — clears the saved session first. Default: false.',
      ),
    },
    async ({ force_new }) => {
      const forceNew = force_new ?? false;

      // Fast path: valid token and not forcing re-login
      if (!forceNew) {
        const existingToken = await getOwaToken();
        if (existingToken) {
          const status = getAuthStatus();
          const mins = status.owaTokenMinutesRemaining ?? 0;

          if (mins >= TOKEN_VALID_THRESHOLD_MINUTES) {
            return {
              content: [{
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  message: `Already authenticated. Token valid for ${mins} more minutes.`,
                  upn: status.upn,
                  tokenStatus: {
                    expiresAt: status.owaTokenExpiry,
                    minutesRemaining: mins,
                  },
                }, null, 2),
              }],
            };
          }
        }
      }

      // Need to open a browser — warn the user not to close it
      if (!forceNew) {
        process.stderr.write(
          '[msoutlook-mcp] Opening browser. Do NOT close the window — it closes automatically once signed in.\n',
        );
      }

      const result = await browserLogin(forceNew);

      if (!result) {
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              success: false,
              message: 'Login failed. If you closed the browser window manually, run outlook_login again and wait for it to close on its own.',
            }, null, 2),
          }],
        };
      }

      // Return path-specific message (mirrors msteams-mcp)
      const messages: Record<typeof result.method, string> = {
        'token-cache':    'Already authenticated. Token valid.',
        'headless-sso':   'Login completed silently via SSO. Session has been saved.',
        'headed-browser': 'Login completed successfully. Session has been saved.',
      };

      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            message: messages[result.method],
            upn: result.upn,
          }, null, 2),
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
      const token = await getOwaToken();

      if (!token) {
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({ authenticated: false, message: 'Not authenticated. Run outlook_login to sign in.' }, null, 2),
          }],
        };
      }

      const status = getAuthStatus();
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            authenticated: true,
            upn: status.upn,
            tenantId: status.tenantId,
            owaToken: {
              expiresAt: status.owaTokenExpiry,
              minutesRemaining: status.owaTokenMinutesRemaining,
            },
            graphToken: status.graphTokenExpiry ? {
              expiresAt: status.graphTokenExpiry,
            } : null,
          }, null, 2),
        }],
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
        content: [{
          type: 'text',
          text: JSON.stringify({ success: true, message: 'Session cleared. Run outlook_login to sign in again.' }, null, 2),
        }],
      };
    },
  );
}
