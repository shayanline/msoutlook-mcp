/**
 * Shared constants used across the codebase.
 */

// ─────────────────────────────────────────────────────────────────────────────
// OAuth / Auth
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Microsoft's own Outlook Web App client ID.
 * This is the first-party app ID used by outlook.office.com itself.
 * No app registration required — same pattern as msteams-mcp uses Teams' own client ID.
 */
export const OWA_CLIENT_ID = '9199bf20-a13f-4107-85dc-02114787ef48';

/** OAuth token endpoint for work/school accounts. */
export const TOKEN_ENDPOINT = 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token';

/** OAuth token endpoint for personal Microsoft accounts. */
export const TOKEN_ENDPOINT_CONSUMERS = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token';

/** Primary resource scope for Outlook Web. */
export const OWA_SCOPE = 'https://outlook.office.com/.default';

/** Microsoft Graph scope. */
export const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

/** Outlook Web login URL. */
export const OWA_URL = 'https://outlook.office.com/mail/';

// ─────────────────────────────────────────────────────────────────────────────
// API Base URLs
// ─────────────────────────────────────────────────────────────────────────────

export const OWA_BASE = 'https://outlook.office.com';
export const OWA_SVC = `${OWA_BASE}/owa/service.svc`;
export const OWA_REST_V2 = `${OWA_BASE}/api/v2.0`;
export const OWA_SEARCH = `${OWA_BASE}/search/api/v1`;
export const OWA_PEOPLE = `${OWA_BASE}/PeopleGraphVx/v1.0`;
export const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

// ─────────────────────────────────────────────────────────────────────────────
// Session Storage
// ─────────────────────────────────────────────────────────────────────────────

export const SESSION_DIR_NAME = '.msoutlook-mcp-server';
export const SESSION_STATE_FILE = 'session-state.json';
export const TOKEN_CACHE_FILE = 'token-cache.json';
export const BROWSER_PROFILE_DIR = 'browser-profile';

// ─────────────────────────────────────────────────────────────────────────────
// Token / Refresh
// ─────────────────────────────────────────────────────────────────────────────

/** Refresh tokens ~55 minutes before expiry (access tokens last 1 hour). */
export const TOKEN_REFRESH_BUFFER_MS = 55 * 60 * 1000;

/** How long to wait for user to complete browser login (ms). */
export const LOGIN_TIMEOUT_MS = 5 * 60 * 1000;

// ─────────────────────────────────────────────────────────────────────────────
// OWA App version header (spoofs the OWA web client)
// ─────────────────────────────────────────────────────────────────────────────

export const OWA_CLIENT_VERSION = '20260522011.06';
export const OWA_USER_AGENT =
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0';

// ─────────────────────────────────────────────────────────────────────────────
// Browser detection
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Ordered list of Playwright channel names to try when auto-detecting a browser.
 * undefined = Playwright's bundled Chromium (always available as final fallback).
 *
 * Override with MSOUTLOOK_BROWSER env var (e.g. "chrome", "msedge", "chromium").
 */
export const BROWSER_CHANNELS: Array<string | undefined> = (() => {
  const override = process.env.MSOUTLOOK_BROWSER?.trim().toLowerCase();
  if (override === 'chromium' || override === 'bundled') return [undefined];
  if (override) return [override, undefined];
  return ['chrome', 'msedge', undefined];
})();
