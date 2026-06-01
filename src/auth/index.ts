/**
 * Auth module — token access, validation, and refresh.
 *
 * Provides `getOwaToken()` and `getGraphToken()` which callers should use
 * instead of reading the cache directly. These functions handle refresh automatically.
 */

import { logger } from '../utils/logger.js';
import { TOKEN_REFRESH_BUFFER_MS } from '../constants.js';
import { readTokenCache } from './session-store.js';
import { refreshOwaToken, refreshGraphToken } from './token-refresh.js';
import { headlessTokenRefresh } from './browser-login.js';

export { browserLogin, headlessTokenRefresh, type LoginResult } from './browser-login.js';
export { clearSession, hasSessionState, isSessionLikelyExpired, readTokenCache } from './session-store.js';

// ─────────────────────────────────────────────────────────────────────────────
// Token Access
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Get a valid OWA access token, refreshing automatically if needed.
 * Returns null if not authenticated or refresh fails.
 */
export async function getOwaToken(): Promise<string | null> {
  const cache = readTokenCache();
  if (!cache) return null;

  // Check if still valid with buffer
  if (cache.owaTokenExpiry - TOKEN_REFRESH_BUFFER_MS > Date.now()) {
    return cache.owaToken;
  }

  logger.debug('OWA token expiring soon, refreshing...');

  // Try HTTP refresh first (fast, no browser)
  const httpRefreshed = await refreshOwaToken();
  if (httpRefreshed) return httpRefreshed;

  // Fall back to headless browser refresh
  logger.info('HTTP refresh failed, attempting headless browser refresh...');
  const browserRefreshed = await headlessTokenRefresh();
  if (browserRefreshed) {
    return readTokenCache()?.owaToken ?? null;
  }

  logger.warn('All refresh methods failed. Run outlook_login to re-authenticate.');
  return null;
}

/**
 * Get a valid Graph access token, refreshing automatically if needed.
 * Falls back to OWA token if Graph token unavailable (some endpoints accept either).
 */
export async function getGraphToken(): Promise<string | null> {
  const cache = readTokenCache();
  if (!cache) return null;

  // Check if Graph token is valid
  if (cache.graphToken && cache.graphTokenExpiry && cache.graphTokenExpiry - TOKEN_REFRESH_BUFFER_MS > Date.now()) {
    return cache.graphToken;
  }

  logger.debug('Graph token expiring soon or missing, refreshing...');

  const refreshed = await refreshGraphToken();
  if (refreshed) return refreshed;

  logger.warn('Graph token refresh failed');
  return null;
}

/**
 * Check if the user is currently authenticated.
 */
export function isAuthenticated(): boolean {
  const cache = readTokenCache();
  if (!cache) return false;
  // Has a valid refresh token (refresh tokens last ~90 days)
  return !!cache.refreshToken;
}

/**
 * Get auth status details for diagnostics.
 */
export function getAuthStatus(): {
  authenticated: boolean;
  upn?: string;
  tenantId?: string;
  owaTokenExpiry?: string;
  owaTokenMinutesRemaining?: number;
  graphTokenExpiry?: string;
} {
  const cache = readTokenCache();
  if (!cache) return { authenticated: false };

  const now = Date.now();
  return {
    authenticated: true,
    upn: cache.upn,
    tenantId: cache.tenantId,
    owaTokenExpiry: new Date(cache.owaTokenExpiry).toISOString(),
    owaTokenMinutesRemaining: Math.max(0, Math.round((cache.owaTokenExpiry - now) / 60_000)),
    graphTokenExpiry: cache.graphTokenExpiry ? new Date(cache.graphTokenExpiry).toISOString() : undefined,
  };
}
