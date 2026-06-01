/**
 * HTTP-based token refresh.
 *
 * Uses the OWA first-party client ID (public SPA — no client secret needed) to
 * exchange a cached refresh token for new access tokens via the standard
 * OAuth2 token endpoint.
 *
 * Key detail (from msteams-mcp): the Origin header is REQUIRED for SPA
 * client IDs. Azure AD validates that refresh token grants from SPA clients
 * include a cross-origin Origin header matching a registered redirect URI.
 * Without it Azure AD returns AADSTS9002327.
 */

import { logger } from '../utils/logger.js';
import { OWA_CLIENT_ID, OWA_SCOPE, GRAPH_SCOPE, OWA_BASE } from '../constants.js';
import { readTokenCache, writeTokenCache, clearTokenCache, type TokenCache } from './session-store.js';

// ─────────────────────────────────────────────────────────────────────────────
// Concurrent refresh guard (mirrors msteams-mcp)
// ─────────────────────────────────────────────────────────────────────────────

let refreshInProgress = false;

// ─────────────────────────────────────────────────────────────────────────────
// Core HTTP call
// ─────────────────────────────────────────────────────────────────────────────

interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  scope?: string;
}

const REFRESH_TIMEOUT_MS = 10_000;

async function callTokenEndpoint(
  tenantId: string,
  refreshToken: string,
  scope: string,
): Promise<TokenResponse | null> {
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const body = new URLSearchParams({
    client_id: OWA_CLIENT_ID,
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    scope,
  });

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), REFRESH_TIMEOUT_MS);

  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        // Required for SPA public clients — Azure AD returns AADSTS9002327 without this
        'Origin': OWA_BASE,
      },
      body: body.toString(),
      signal: controller.signal,
    });

    if (!res.ok) {
      const text = await res.text().catch(() => '');
      logger.warn(`Token refresh HTTP ${res.status}`, text.slice(0, 200));
      return null;
    }

    return res.json() as Promise<TokenResponse>;
  } catch (err) {
    if (err instanceof Error && err.name === 'AbortError') {
      logger.warn('Token refresh request timed out');
    } else {
      logger.warn('Token refresh network error', err instanceof Error ? err.message : String(err));
    }
    return null;
  } finally {
    clearTimeout(timer);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Public refresh functions
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Refresh the OWA access token via HTTP.
 * Returns the new access token on success, null on failure.
 * Prevents concurrent refresh races with module-level guard.
 */
export async function refreshOwaToken(): Promise<string | null> {
  if (refreshInProgress) {
    logger.debug('Token refresh already in progress — waiting is not needed, caller should retry');
    return null;
  }

  const cache = readTokenCache();
  if (!cache?.refreshToken || !cache.tenantId) {
    logger.debug('No refresh token or tenant ID cached');
    return null;
  }

  refreshInProgress = true;
  try {
    logger.debug('Refreshing OWA access token via HTTP');
    const response = await callTokenEndpoint(cache.tenantId, cache.refreshToken, OWA_SCOPE);
    if (!response) return null;

    const updated: TokenCache = {
      ...cache,
      owaToken: response.access_token,
      owaTokenExpiry: Date.now() + response.expires_in * 1000,
      refreshToken: response.refresh_token ?? cache.refreshToken,
      extractedAt: Date.now(),
    };
    writeTokenCache(updated);
    logger.info('OWA token refreshed successfully');
    return response.access_token;
  } finally {
    refreshInProgress = false;
  }
}

/**
 * Refresh the Graph access token via HTTP.
 */
export async function refreshGraphToken(): Promise<string | null> {
  if (refreshInProgress) return null;

  const cache = readTokenCache();
  if (!cache?.refreshToken || !cache.tenantId) return null;

  refreshInProgress = true;
  try {
    logger.debug('Refreshing Graph access token via HTTP');
    const response = await callTokenEndpoint(cache.tenantId, cache.refreshToken, GRAPH_SCOPE);
    if (!response) return null;

    const updated: TokenCache = {
      ...cache,
      graphToken: response.access_token,
      graphTokenExpiry: Date.now() + response.expires_in * 1000,
      refreshToken: response.refresh_token ?? cache.refreshToken,
      extractedAt: Date.now(),
    };
    writeTokenCache(updated);
    logger.info('Graph token refreshed successfully');
    return response.access_token;
  } finally {
    refreshInProgress = false;
  }
}
