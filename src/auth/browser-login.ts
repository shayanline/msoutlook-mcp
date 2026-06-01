/**
 * Playwright-based browser login flow for Outlook Web App.
 *
 * Ported closely from msteams-mcp's browser auth approach:
 * - Uses the system default browser (Edge on Windows, Chrome on macOS)
 * - Imports Microsoft SSO cookies from the user's real browser profile,
 *   enabling silent authentication without typing credentials
 * - Headless first, visible browser only as last resort
 * - Handles stale SingletonLock files from crashed sessions
 */

import { chromium, type BrowserContext } from 'playwright';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { execSync } from 'node:child_process';
import { logger } from '../utils/logger.js';
import { OWA_URL, LOGIN_TIMEOUT_MS } from '../constants.js';
import {
  getBrowserProfileDir,
  writeSessionState,
  writeTokenCache,
  type TokenCache,
} from './session-store.js';
import { extractTokensFromLocalStorage, getOwaLocalStorage } from './token-extractor.js';
import { importMicrosoftCookies } from '../browser/cookie-import.js';

// ─────────────────────────────────────────────────────────────────────────────
// Browser channel
// ─────────────────────────────────────────────────────────────────────────────

/** Microsoft login domains — redirect to these means we're not authenticated. */
const LOGIN_DOMAINS = [
  'login.microsoftonline.com',
  'login.live.com',
  'login.microsoft.com',
];

/**
 * Determine the Playwright channel to use.
 *
 * Priority:
 * 1. MSOUTLOOK_BROWSER env var (explicit override)
 * 2. macOS system default browser (read from LaunchServices)
 * 3. Platform fallback: Chrome on macOS/Linux, Edge on Windows
 */
function getBrowserChannel(): string {
  const override = process.env.MSOUTLOOK_BROWSER?.trim().toLowerCase();
  if (override && override !== 'chromium' && override !== 'bundled') return override;

  if (process.platform === 'darwin') {
    const detected = getMacOSDefaultBrowser();
    if (detected) {
      logger.debug(`macOS default browser detected: ${detected}`);
      return detected;
    }
    return 'chrome'; // macOS fallback (same as msteams-mcp)
  }

  return process.platform === 'win32' ? 'msedge' : 'chrome';
}

/**
 * Detect the macOS default browser by reading the LaunchServices plist.
 * Returns a Playwright channel name or undefined if unrecognised.
 */
function getMacOSDefaultBrowser(): string | undefined {
  try {
    const plistPath = path.join(
      process.env.HOME ?? '',
      'Library/Preferences/com.apple.LaunchServices/com.apple.launchservices.secure.plist',
    );
    const json = execSync(`plutil -convert json -o - "${plistPath}"`, {
      encoding: 'utf8',
      timeout: 3000,
    });
    const data = JSON.parse(json) as { LSHandlers?: Array<{ LSHandlerURLScheme?: string; LSHandlerRoleAll?: string }> };
    const handlers = data.LSHandlers ?? [];
    const httpsHandler = handlers.find(h => h.LSHandlerURLScheme === 'https');
    const bundleId = (httpsHandler?.LSHandlerRoleAll ?? '').toLowerCase();

    if (bundleId.includes('microsoft.edgemac') || bundleId.includes('edge')) return 'msedge';
    if (bundleId.includes('google.chrome') || bundleId.includes('chrome')) return 'chrome';
    // Firefox and Safari don't have Playwright channels — fall through to platform default
    return undefined;
  } catch {
    return undefined;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// SingletonLock cleanup (same as msteams-mcp)
// ─────────────────────────────────────────────────────────────────────────────

function cleanupStaleSingletonLock(profileDir: string): void {
  const lockPath = path.join(profileDir, 'SingletonLock');
  if (!fs.existsSync(lockPath)) return;

  try {
    const linkTarget = fs.readlinkSync(lockPath);
    const match = linkTarget.match(/-(\d+)$/);
    if (match) {
      const pid = parseInt(match[1], 10);
      try {
        process.kill(pid, 0); // signal 0 = check if process exists
        return; // still running — don't remove
      } catch {
        // process is gone — remove stale lock
      }
    }
    fs.unlinkSync(lockPath);
    logger.debug('Removed stale SingletonLock');
  } catch {
    // ignore
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Shared helpers
// ─────────────────────────────────────────────────────────────────────────────

function hasSavedBrowserProfile(): boolean {
  const dir = getBrowserProfileDir();
  return fs.existsSync(dir) && fs.readdirSync(dir).length > 0;
}

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

/**
 * Launch a persistent context, clean up stale locks first.
 * Falls back to removing the lock and retrying once if a lock conflict occurs.
 */
async function launchContext(profileDir: string, headless: boolean, channel: string): Promise<BrowserContext> {
  cleanupStaleSingletonLock(profileDir);

  const launch = () => chromium.launchPersistentContext(profileDir, {
    headless,
    channel,
    viewport: { width: 1280, height: 800 },
    acceptDownloads: false,
  });

  try {
    return await launch();
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    if (msg.includes('ProcessSingleton') || msg.includes('SingletonLock')) {
      logger.debug('Profile lock conflict — removing lock and retrying');
      const lockPath = path.join(profileDir, 'SingletonLock');
      try { fs.unlinkSync(lockPath); } catch { /* ignore */ }
      return launch();
    }
    throw err;
  }
}

/**
 * Wait for OWA to authenticate AND for MSAL tokens to be present in localStorage.
 * Uses a redirect-detection approach (like msteams-mcp) for fast failure on
 * unauthenticated sessions, so we don't sit at a blank login page for 30s+.
 *
 * @returns true if authenticated with tokens, false if redirected to login
 */
async function waitForOwaAuth(context: BrowserContext, timeoutMs: number): Promise<boolean> {
  const page = context.pages()[0] ?? await context.newPage();

  // Detect redirect to Microsoft login — fast signal that the session is invalid
  let redirectedToLogin = false;
  const onNavigation = (frame: { url: () => string }) => {
    if (frame === page.mainFrame() && LOGIN_DOMAINS.some(d => frame.url().includes(d))) {
      redirectedToLogin = true;
    }
  };
  page.on('framenavigated', onNavigation);

  try {
    await page.goto(OWA_URL, { waitUntil: 'domcontentloaded' });

    // Give MSAL 5 s to either redirect to login or silently authenticate.
    // If redirected to login page, return false immediately.
    const deadline = Date.now() + Math.min(timeoutMs, 5_000);
    while (Date.now() < deadline) {
      if (redirectedToLogin) return false;
      await page.waitForTimeout(100);
    }
    if (redirectedToLogin) return false;

    // No login redirect — session appears valid. Now wait for MSAL tokens.
    await page.waitForFunction(
      () => {
        if (localStorage.getItem('olk-isauthed') !== 'true') return false;
        return Object.keys(localStorage).some(
          k => k.includes('|accesstoken|') && k.includes('outlook.office.com'),
        );
      },
      { timeout: timeoutMs },
    );
    return true;
  } finally {
    page.off('framenavigated', onNavigation);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Headless login (primary)
// ─────────────────────────────────────────────────────────────────────────────

export async function headlessLogin(): Promise<string | null> {
  logger.debug('Attempting headless (silent) login...');

  const profileDir = getBrowserProfileDir();
  const channel = getBrowserChannel();
  let context: BrowserContext | null = null;

  try {
    context = await launchContext(profileDir, true, channel);

    const authenticated = await waitForOwaAuth(context, 15_000);
    if (!authenticated) {
      logger.debug('Headless: session expired or not present');
      return null;
    }

    const upn = await extractAndCacheTokens(context);
    if (upn) logger.info(`Headless login succeeded (${upn})`);
    return upn;
  } catch {
    logger.debug('Headless login failed');
    return null;
  } finally {
    await context?.close().catch(() => {});
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Headed login (fallback)
// ─────────────────────────────────────────────────────────────────────────────

export async function headedLogin(): Promise<string | null> {
  logger.info('Opening browser for interactive login...');

  const profileDir = getBrowserProfileDir();
  const channel = getBrowserChannel();
  let context: BrowserContext | null = null;
  let browserClosed = false;

  try {
    context = await launchContext(profileDir, false, channel);
    context.on('close', () => { browserClosed = true; });

    // Import Microsoft SSO cookies from the user's real browser — enables
    // instant silent sign-in without typing credentials (same as msteams-mcp).
    await importMicrosoftCookies(context, channel);

    const authenticated = await waitForOwaAuth(context, LOGIN_TIMEOUT_MS);

    if (!authenticated) {
      // Login redirect — user needs to sign in manually; wait for them to complete
      logger.info('Waiting for you to complete sign-in in the browser...');
      const page = context.pages()[0];
      if (page) {
        await page.waitForFunction(
          () => {
            if (localStorage.getItem('olk-isauthed') !== 'true') return false;
            return Object.keys(localStorage).some(
              k => k.includes('|accesstoken|') && k.includes('outlook.office.com'),
            );
          },
          { timeout: LOGIN_TIMEOUT_MS },
        );
      }
    }

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

export async function browserLogin(): Promise<string | null> {
  if (hasSavedBrowserProfile()) {
    const upn = await headlessLogin();
    if (upn) return upn;
    logger.info('Headless login failed — falling back to visible browser...');
  } else {
    logger.info('No saved browser profile — opening visible browser for first-time setup...');
  }
  return headedLogin();
}

// ─────────────────────────────────────────────────────────────────────────────
// Token refresh alias
// ─────────────────────────────────────────────────────────────────────────────

export async function headlessTokenRefresh(): Promise<boolean> {
  return (await headlessLogin()) !== null;
}
