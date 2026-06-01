/**
 * HTTP-based token refresh using the MSAL refresh token.
 *
 * Uses the OWA first-party client ID (no client secret needed — OWA is a public client SPA)
 * to exchange a refresh token for new access tokens via the standard OAuth2 token endpoint.
 */

import { logger } from '../utils/logger.js';
import {
  OWA_CLIENT_ID,
  OWA_SCOPE,
  GRAPH_SCOPE,
  TOKEN_ENDPOINT,
  TOKEN_ENDPOINT_CONSUMERS,
} from '../constants.js';
import { readTokenCache, writeTokenCache, type TokenCache } from './session-store.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  token_type: string;
  scope?: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// Refresh
// ─────────────────────────────────────────────────────────────────────────────

async function refreshTokenForScope(
  refreshToken: string,
  scope: string,
  tenantId?: string,
): Promise<TokenResponse | null> {
  const endpoint = tenantId
    ? `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`
    : TOKEN_ENDPOINT;

  const body = new URLSearchParams({
    client_id: OWA_CLIENT_ID,
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    scope,
  });

  try {
    const res = await fetch(endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body.toString(),
    });

    if (!res.ok) {
      const text = await res.text();
      logger.warn(`Token refresh failed (${res.status})`, text.slice(0, 200));
      return null;
    }

    return res.json() as Promise<TokenResponse>;
  } catch (err) {
    logger.warn('Token refresh network error', err);
    return null;
  }
}

/**
 * Refresh OWA access token using the cached refresh token.
 * Updates the token cache on success.
 * Returns the new access token or null if refresh failed.
 */
export async function refreshOwaToken(): Promise<string | null> {
  const cache = readTokenCache();
  if (!cache?.refreshToken) {
    logger.debug('No refresh token cached, cannot refresh');
    return null;
  }

  logger.debug('Refreshing OWA access token via HTTP');
  const response = await refreshTokenForScope(cache.refreshToken, OWA_SCOPE, cache.tenantId);
  if (!response) return null;

  const expiry = Date.now() + response.expires_in * 1000;
  const updated: TokenCache = {
    ...cache,
    owaToken: response.access_token,
    owaTokenExpiry: expiry,
    // Some flows return a new refresh token; use it if provided
    refreshToken: response.refresh_token ?? cache.refreshToken,
    extractedAt: Date.now(),
  };

  writeTokenCache(updated);
  logger.info('OWA token refreshed successfully');
  return response.access_token;
}

/**
 * Refresh Graph access token using the cached refresh token.
 */
export async function refreshGraphToken(): Promise<string | null> {
  const cache = readTokenCache();
  if (!cache?.refreshToken) return null;

  logger.debug('Refreshing Graph access token via HTTP');
  const response = await refreshTokenForScope(cache.refreshToken, GRAPH_SCOPE, cache.tenantId);
  if (!response) return null;

  const expiry = Date.now() + response.expires_in * 1000;
  const updated: TokenCache = {
    ...cache,
    graphToken: response.access_token,
    graphTokenExpiry: expiry,
    refreshToken: response.refresh_token ?? cache.refreshToken,
    extractedAt: Date.now(),
  };

  writeTokenCache(updated);
  logger.info('Graph token refreshed successfully');
  return response.access_token;
}
