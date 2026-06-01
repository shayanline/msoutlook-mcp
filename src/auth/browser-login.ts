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
import { existsSync, readdirSync } from 'node:fs';
import { logger } from '../utils/logger.js';
import { OWA_URL, LOGIN_TIMEOUT_MS, BROWSER_CHANNELS } from '../constants.js';
import {
  getBrowserProfileDir,
  writeSessionState,
  writeTokenCache,
  type TokenCache,
} from './session-store.js';
import { extractTokensFromLocalStorage, getOwaLocalStorage } from './token-extractor.js';

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Returns true if the browser profile directory exists and has been populated
 * by a previous login. Empty or missing profile = headless will definitely fail.
 */
function hasSavedBrowserProfile(): boolean {
  const dir = getBrowserProfileDir();
  return existsSync(dir) && readdirSync(dir).length > 0;
}

/** Cached result of browser detection so we only probe once per process. */
let detectedChannel: string | undefined | null = null; // null = not yet detected

/**
 * Detect the best available Chromium-based browser on this machine.
 * Tries each channel in BROWSER_CHANNELS order; the first one that launches
 * successfully wins. Falls back to Playwright's bundled Chromium (undefined).
 */
async function detectBrowserChannel(): Promise<string | undefined> {
  if (detectedChannel !== null) return detectedChannel;

  for (const channel of BROWSER_CHANNELS) {
    try {
      const browser = await chromium.launch({ channel, headless: true });
      await browser.close();
      const label = channel ?? 'bundled Chromium';
      logger.info(`Using browser: ${label}`);
      detectedChannel = channel;
      return channel;
    } catch {
      continue;
    }
  }

  // Should never reach here — bundled Chromium (undefined) always works
  detectedChannel = undefined;
  return undefined;
}

/** Wait for OWA to fully authenticate AND for MSAL to populate tokens in localStorage. */
async function waitForOwaAuth(context: BrowserContext, timeoutMs: number): Promise<void> {
  const page = context.pages()[0] ?? await context.newPage();
  await page.goto(OWA_URL, { waitUntil: 'domcontentloaded' });

  // Two conditions must both be true before we proceed:
  // 1. OWA sets olk-isauthed once the app shell is ready
  // 2. At least one OWA-scoped MSAL access token is present in localStorage
  //
  // Checking for the token directly avoids the race condition where the URL or
  // olk-isauthed flag is set but MSAL hasn't finished its silent token acquisition.
  await page.waitForFunction(
    () => {
      if (localStorage.getItem('olk-isauthed') !== 'true') return false;
      return Object.keys(localStorage).some(
        k => k.includes('|accesstoken|') && k.includes('outlook.office.com'),
      );
    },
    { timeout: timeoutMs },
  );
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
    const channel = await detectBrowserChannel();
    context = await chromium.launchPersistentContext(profileDir, {
      headless: true,
      channel,
    });

    // 10 s is plenty — if the saved session is valid it loads in 2–4 s.
    // Failing fast avoids a 30 s wait before falling back to headed.
    await waitForOwaAuth(context, 10_000);

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
    const channel = await detectBrowserChannel();
    context = await chromium.launchPersistentContext(profileDir, {
      headless: false,
      channel,
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
  // 1. Try headless only if a saved profile exists — pointless otherwise
  if (hasSavedBrowserProfile()) {
    const upn = await headlessLogin();
    if (upn) return upn;
    logger.info('Headless login failed. Falling back to visible browser...');
  } else {
    logger.info('No saved browser profile — going straight to visible login.');
  }

  // 2. Headed fallback — opens Edge so the user can sign in interactively
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
