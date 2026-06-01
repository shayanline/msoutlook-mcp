/**
 * Extract MSAL tokens from Outlook Web App's localStorage.
 *
 * OWA uses MSAL v3 and stores access tokens, refresh tokens, and ID tokens
 * in localStorage under keys like:
 *   msal.3|{accountId}|login.windows.net|accesstoken|{clientId}|{tenantId}|{scopes}
 *   msal.3|{accountId}|login.windows.net|refreshtoken|{clientId}|||
 *
 * This mirrors how msteams-mcp extracts Teams tokens, adapted for OWA.
 */

import { logger } from '../utils/logger.js';
import { OWA_CLIENT_ID } from '../constants.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

export interface ExtractedTokens {
  owaToken: string;
  owaTokenExpiry: Date;
  graphToken?: string;
  graphTokenExpiry?: Date;
  refreshToken: string;
  tenantId?: string;
  upn?: string;
}

interface MsalEntry {
  secret: string;
  credentialType?: string;
  target?: string;
  realm?: string;
  homeAccountId?: string;
  clientId?: string;
  environment?: string;
}

interface JwtPayload {
  exp?: number;
  aud?: string;
  upn?: string;
  preferred_username?: string;
  tid?: string;
  appid?: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// JWT utilities
// ─────────────────────────────────────────────────────────────────────────────

function decodeJwtPayload(token: string): JwtPayload | null {
  try {
    const parts = token.split('.');
    if (parts.length < 2) return null;
    // Handle base64url encoding
    const b64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
    const padded = b64 + '='.repeat((4 - (b64.length % 4)) % 4);
    return JSON.parse(Buffer.from(padded, 'base64').toString('utf8')) as JwtPayload;
  } catch {
    return null;
  }
}

function getJwtExpiry(token: string): Date | null {
  const payload = decodeJwtPayload(token);
  if (!payload?.exp || typeof payload.exp !== 'number') return null;
  return new Date(payload.exp * 1000);
}

function isJwt(value: string): boolean {
  return typeof value === 'string' && value.startsWith('ey');
}

// ─────────────────────────────────────────────────────────────────────────────
// Extraction from Playwright storageState localStorage entries
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Extract all MSAL tokens from the Playwright storageState localStorage array.
 *
 * @param localStorage Array of {name, value} entries from Playwright storageState
 */
export function extractTokensFromLocalStorage(
  localStorage: Array<{ name: string; value: string }>,
): ExtractedTokens | null {
  let bestOwaToken: { token: string; expiry: Date } | null = null;
  let bestGraphToken: { token: string; expiry: Date } | null = null;
  let refreshToken: string | null = null;
  let tenantId: string | undefined;
  let upn: string | undefined;

  for (const item of localStorage) {
    const key = item.name;
    if (!key.startsWith('msal.')) continue;

    let entry: MsalEntry;
    try {
      entry = JSON.parse(item.value) as MsalEntry;
    } catch {
      continue;
    }

    if (!entry.secret) continue;

    // ── Refresh Token ─────────────────────────────────────────────────────
    if (key.includes('refreshtoken') && entry.clientId === OWA_CLIENT_ID) {
      refreshToken = entry.secret;
      continue;
    }

    // ── Access Tokens ─────────────────────────────────────────────────────
    if (!key.includes('accesstoken')) continue;
    if (!isJwt(entry.secret)) continue;

    const expiry = getJwtExpiry(entry.secret);
    if (!expiry) continue;
    if (expiry.getTime() <= Date.now()) continue; // skip expired

    // Extract UPN and tenant from any token
    if (!upn) {
      const payload = decodeJwtPayload(entry.secret);
      if (payload) {
        upn = payload.upn ?? payload.preferred_username;
        tenantId = tenantId ?? payload.tid;
      }
    }

    // OWA token — scope contains outlook.office.com
    if (entry.target?.includes('outlook.office.com')) {
      if (!bestOwaToken || expiry > bestOwaToken.expiry) {
        bestOwaToken = { token: entry.secret, expiry };
      }
    }

    // Graph token — scope contains graph.microsoft.com
    if (entry.target?.includes('graph.microsoft.com')) {
      if (!bestGraphToken || expiry > bestGraphToken.expiry) {
        bestGraphToken = { token: entry.secret, expiry };
      }
    }
  }

  if (!bestOwaToken || !refreshToken) {
    logger.debug('Token extraction failed', {
      hasOwaToken: !!bestOwaToken,
      hasRefreshToken: !!refreshToken,
    });
    return null;
  }

  return {
    owaToken: bestOwaToken.token,
    owaTokenExpiry: bestOwaToken.expiry,
    graphToken: bestGraphToken?.token,
    graphTokenExpiry: bestGraphToken?.expiry,
    refreshToken,
    tenantId,
    upn,
  };
}

/** Shape of the relevant parts of a Playwright storageState. */
export interface StorageState {
  origins?: Array<{
    origin: string;
    localStorage: Array<{ name: string; value: string }>;
  }>;
}

/**
 * Get the localStorage array from a Playwright storageState for the OWA origin.
 */
export function getOwaLocalStorage(
  state: StorageState,
): Array<{ name: string; value: string }> | null {
  const owaOrigin = state.origins?.find(o => o.origin.includes('outlook.office.com'));
  return owaOrigin?.localStorage ?? null;
}
