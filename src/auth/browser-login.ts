/**
 * Playwright-based browser login flow for Outlook Web App.
 *
 * Priority order:
 * 1. Headless (silent) — reuses the persisted browser profile; user sees nothing
 * 2. Headed (visible) — fallback when the session has expired and the user must sign in
 *
 * After the initial headed login the session is persisted, so subsequent logins
 * go through the headless path silently.
 */

import { chromium, type BrowserContext } from 'playwright';
import { logger } from '../utils/logger.js';
import { OWA_URL, LOGIN_TIMEOUT_MS } from '../constants.js';
import {
  getBrowserProfileDir,
  writeSessionState,
  writeTokenCache,
  type TokenCache,
} from './session-store.js';
import { extractTokensFromLocalStorage, getOwaLocalStorage } from './token-extractor.js';

// ─────────────────────────────────────────────────────────────────────────────
// Shared helpers
// ─────────────────────────────────────────────────────────────────────────────

/** Wait for OWA to report itself as authenticated. */
async function waitForOwaAuth(context: BrowserContext, timeoutMs: number): Promise<void> {
  const page = context.pages()[0] ?? await context.newPage();
  await page.goto(OWA_URL, { waitUntil: 'domcontentloaded' });

  await Promise.race([
    page.waitForFunction(
      () => localStorage.getItem('olk-isauthed') === 'true',
      { timeout: timeoutMs },
    ),
    page.waitForURL('**/mail/**', { timeout: timeoutMs }),
  ]);
}

/** Extract MSAL tokens from the live browser context and write to cache. */
async function extractAndCacheTokens(context: BrowserContext): Promise<string | null> {
  const state = await context.storageState();
  writeSessionState(state);

  const ls = getOwaLocalStorage(state as unknown as Record<string, unknown>);
  if (!ls) {
    logger.debug('Could not find Outlook origin in session state');
    return null;
  }

  const tokens = extractTokensFromLocalStorage(ls);
  if (!tokens) {
    logger.debug('Could not extract OWA tokens from localStorage');
    return null;
  }

  const cache: TokenCache = {
    owaToken: tokens.owaToken,
    owaTokenExpiry: tokens.owaTokenExpiry.getTime(),
    graphToken: tokens.graphToken,
    graphTokenExpiry: tokens.graphTokenExpiry?.getTime(),
    refreshToken: tokens.refreshToken,
    tenantId: tokens.tenantId,
    upn: tokens.upn,
    extractedAt: Date.now(),
  };
  writeTokenCache(cache);
  return tokens.upn ?? 'unknown';
}

// ─────────────────────────────────────────────────────────────────────────────
// Headless login (primary path)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Attempt a completely silent login using the persisted browser profile.
 * If the saved session cookies are still valid, MSAL will silently restore
 * the session and we can extract fresh tokens without showing anything to the user.
 *
 * @returns the signed-in UPN on success, null if the session has expired
 */
export async function headlessLogin(): Promise<string | null> {
  logger.debug('Attempting headless (silent) login...');

  const profileDir = getBrowserProfileDir();
  let context: BrowserContext | null = null;

  try {
    context = await chromium.launchPersistentContext(profileDir, {
      headless: true,
      channel: 'msedge',
    });

    await waitForOwaAuth(context, 30_000);

    const upn = await extractAndCacheTokens(context);
    if (upn) logger.info(`Headless login succeeded (${upn})`);
    return upn;
  } catch {
    // Session expired or not present — caller should fall back to headed login
    logger.debug('Headless login failed — session likely expired');
    return null;
  } finally {
    await context?.close().catch(() => {});
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Headed login (fallback)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Open a visible browser window so the user can sign in interactively.
 * Used only when headless login fails (session expired / first-time setup).
 * The browser closes itself automatically once authentication completes.
 *
 * @returns the signed-in UPN on success, null on failure or manual close
 */
export async function headedLogin(): Promise<string | null> {
  logger.info('Opening browser for interactive login (session expired or first-time setup)...');

  const profileDir = getBrowserProfileDir();
  let context: BrowserContext | null = null;
  let browserClosed = false;

  try {
    context = await chromium.launchPersistentContext(profileDir, {
      headless: false,
      channel: 'msedge',
      args: ['--no-sandbox', '--disable-dev-shm-usage'],
    });

    context.on('close', () => { browserClosed = true; });

    await waitForOwaAuth(context, LOGIN_TIMEOUT_MS);

    const upn = await extractAndCacheTokens(context);
    if (upn) logger.info(`Headed login succeeded (${upn}). Browser closing.`);
    return upn;
  } catch (err) {
    if (browserClosed) {
      logger.error('Login aborted — browser was closed before authentication completed.');
    } else {
      logger.error('Headed login failed', err instanceof Error ? err.message : String(err));
    }
    return null;
  } finally {
    if (context && !browserClosed) {
      await context.close().catch(() => {});
    }
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Main entry point
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Login to Outlook Web, preferring headless (silent) mode.
 * Falls back to a visible browser window only if the session has expired.
 *
 * @returns the signed-in UPN on success, null on failure
 */
export async function browserLogin(): Promise<string | null> {
  // 1. Try headless first — completely silent if session is still alive
  const upn = await headlessLogin();
  if (upn) return upn;

  // 2. Session expired — open a visible browser so the user can sign in
  logger.info('Headless login failed. Falling back to visible browser...');
  return headedLogin();
}

// ─────────────────────────────────────────────────────────────────────────────
// Token refresh (used by auth/index.ts)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Attempt a silent browser-based token refresh.
 * Alias for headlessLogin() used in the token refresh pipeline.
 */
export async function headlessTokenRefresh(): Promise<boolean> {
  const upn = await headlessLogin();
  return upn !== null;
}
