/**
 * Playwright-based browser login flow for Outlook Web App.
 *
 * Opens a browser window to outlook.office.com, waits for the user to
 * sign in (or reuses an existing session), then extracts the MSAL tokens
 * from localStorage and caches them.
 *
 * Mirrors the approach used by msteams-mcp for Microsoft Teams.
 */

import { chromium, type BrowserContext } from 'playwright';
import { logger } from '../utils/logger.js';
import {
  OWA_URL,
  LOGIN_TIMEOUT_MS,
  BROWSER_PROFILE_DIR,
} from '../constants.js';
import {
  getBrowserProfileDir,
  writeSessionState,
  writeTokenCache,
  type TokenCache,
} from './session-store.js';
import {
  extractTokensFromLocalStorage,
  getOwaLocalStorage,
} from './token-extractor.js';

// ─────────────────────────────────────────────────────────────────────────────
// Login
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Open a browser to Outlook Web and wait for successful login.
 * Extracts and caches tokens on success.
 *
 * @returns true if login succeeded and tokens were extracted
 */
export async function browserLogin(): Promise<boolean> {
  logger.info('Opening browser for Outlook Web login...');

  const profileDir = getBrowserProfileDir();
  let context: BrowserContext | null = null;

  try {
    context = await chromium.launchPersistentContext(profileDir, {
      headless: false,
      channel: 'msedge',
      args: [
        '--no-sandbox',
        '--disable-dev-shm-usage',
      ],
    });

    const page = context.pages()[0] ?? await context.newPage();
    await page.goto(OWA_URL, { waitUntil: 'domcontentloaded' });

    logger.info('Waiting for Outlook to load and authenticate...');

    // Wait until the mail app is loaded (the olk-isauthed flag is set)
    await page.waitForFunction(
      () => localStorage.getItem('olk-isauthed') === 'true',
      { timeout: LOGIN_TIMEOUT_MS },
    );

    logger.info('Outlook authenticated. Extracting tokens...');

    // Save Playwright storage state (localStorage + cookies)
    const state = await context.storageState();
    writeSessionState(state);

    // Extract tokens from localStorage
    const ls = getOwaLocalStorage(state as unknown as Record<string, unknown>);
    if (!ls) {
      logger.error('Could not find Outlook origin in session state');
      return false;
    }

    const tokens = extractTokensFromLocalStorage(ls);
    if (!tokens) {
      logger.error('Could not extract OWA tokens from localStorage');
      return false;
    }

    // Cache the extracted tokens
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

    logger.info(`Logged in as ${tokens.upn ?? 'unknown'}. Tokens cached.`);
    return true;
  } catch (err) {
    logger.error('Browser login failed', err);
    return false;
  } finally {
    if (context) {
      await context.close().catch(() => {});
    }
  }
}

/**
 * Attempt a headless token refresh using the persisted browser profile.
 * This opens a headless browser to outlook.office.com — if the session
 * cookie is still valid, MSAL will silently refresh the access token.
 *
 * @returns true if tokens were refreshed successfully
 */
export async function headlessTokenRefresh(): Promise<boolean> {
  logger.debug('Attempting headless token refresh via browser...');

  const profileDir = getBrowserProfileDir();
  let context: BrowserContext | null = null;

  try {
    context = await chromium.launchPersistentContext(profileDir, {
      headless: true,
      channel: 'msedge',
    });

    const page = context.pages()[0] ?? await context.newPage();
    await page.goto(OWA_URL, { waitUntil: 'domcontentloaded' });

    // Wait briefly for MSAL to run its silent token acquisition
    await page.waitForFunction(
      () => localStorage.getItem('olk-isauthed') === 'true',
      { timeout: 30_000 },
    );

    const state = await context.storageState();
    writeSessionState(state);

    const ls = getOwaLocalStorage(state as unknown as Record<string, unknown>);
    if (!ls) return false;

    const tokens = extractTokensFromLocalStorage(ls);
    if (!tokens) return false;

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

    logger.info('Headless token refresh succeeded');
    return true;
  } catch (err) {
    logger.debug('Headless token refresh failed', err);
    return false;
  } finally {
    if (context) {
      await context.close().catch(() => {});
    }
  }
}
