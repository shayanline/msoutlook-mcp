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

/** OAuth token endpoint for work/school accounts (used as default when no tenant ID). */
export const TOKEN_ENDPOINT = 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token';

/** Primary resource scope for Outlook Web. */
export const OWA_SCOPE = 'https://outlook.office.com/.default';

/** Microsoft Graph scope. */
export const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

/**
 * Outlook Web login URL.
 *
 * Microsoft is unifying authenticated Microsoft 365 apps under cloud.microsoft
 * (see MC950871 / https://learn.microsoft.com/microsoft-365/enterprise/cloud-microsoft-domain).
 * Outlook on the web is served from outlook.cloud.microsoft; legacy hosts such as
 * outlook.office.com remain supported and redirect during the transition.
 */
export const OWA_URL = 'https://outlook.cloud.microsoft/mail/';

// ─────────────────────────────────────────────────────────────────────────────
// API Base URLs
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Outlook REST API host, distinct from the web UI host (`OWA_URL`): cloud.microsoft is
 * the authenticated app experience; REST + token audience remain `outlook.office.com`
 * (see Outlook REST docs / `https://outlook.office.com/.default`).
 */
export const OWA_BASE = 'https://outlook.office.com';
export const OWA_REST_V2 = `${OWA_BASE}/api/v2.0`;
export const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

// ─────────────────────────────────────────────────────────────────────────────
// Token / Refresh
// ─────────────────────────────────────────────────────────────────────────────

/** Refresh tokens ~55 minutes before expiry (access tokens last 1 hour). */
export const TOKEN_REFRESH_BUFFER_MS = 55 * 60 * 1000;

/** How long to wait for user to complete browser login (ms). */
export const LOGIN_TIMEOUT_MS = 5 * 60 * 1000;

// ─────────────────────────────────────────────────────────────────────────────
// HTTP
// ─────────────────────────────────────────────────────────────────────────────

/** User-Agent presented to OWA APIs so requests look like the web client. */
export const OWA_USER_AGENT =
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0';
