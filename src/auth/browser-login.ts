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
  clearSession,
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
 * Predicate evaluated INSIDE the browser. Returns true once an OWA-scoped MSAL
 * access token is present in localStorage. This is the real signal that auth
 * has completed — we deliberately do NOT check the `olk-isauthed` flag, whose
 * value is "1" (not "true") and varies, while this token check is what the
 * token extractor itself relies on.
 */
function owaAccessTokenPresent(): boolean {
  return Object.keys(localStorage).some(
    k => k.includes('accesstoken') && k.includes('outlook.office.com'),
  );
}

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

  const ls = getOwaLocalStorage(state);
  if (!ls) {
    logger.debug('Could not find Outlook origin in session state');
    return null;
  }

  // Pass cookies so MSAL v4 encrypted localStorage entries can be decrypted
  // via the msal.cache.encryption session cookie.
  const tokens = await extractTokensFromLocalStorage(ls, state.cookies);
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

    // No login redirect — session appears valid. Now wait for the MSAL token.
    await page.waitForFunction(owaAccessTokenPresent, { timeout: timeoutMs });
    return true;
  } finally {
    page.off('framenavigated', onNavigation);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// MSAL token polling (mirrors msteams-mcp waitForTokenRefresh)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * When the browser session is valid (no login redirect) but MSAL tokens are
 * still being acquired, poll localStorage in-browser until tokens appear.
 * This is the case where session cookies are alive but the hour-long access
 * token has expired — OWA JS silently acquires new ones.
 *
 * Polls in-browser to avoid deserialising the full session state on every
 * check. Returns the UPN once tokens are available.
 */
async function waitForMsalTokens(
  context: BrowserContext,
  timeoutMs = 20_000,
): Promise<boolean> {
  const page = context.pages()[0];
  if (!page) return false;

  logger.debug('Waiting for MSAL to acquire tokens...');
  const interval = 1_000;
  const deadline = Date.now() + timeoutMs;

  while (Date.now() < deadline) {
    const ready = await page.evaluate(owaAccessTokenPresent).catch(() => false);

    if (ready) {
      logger.debug(`MSAL tokens appeared after ${Math.round((Date.now() - (deadline - timeoutMs)) / 1000)}s`);
      return true;
    }
    await page.waitForTimeout(interval);
  }

  logger.debug('MSAL token wait timed out');
  return false;
}

// ─────────────────────────────────────────────────────────────────────────────
// Headless login (primary)
// ─────────────────────────────────────────────────────────────────────────────

export async function headlessLogin(): Promise<LoginResult | null> {
  logger.debug('Attempting headless (silent) login...');

  const profileDir = getBrowserProfileDir();
  const channel = getBrowserChannel();
  let context: BrowserContext | null = null;

  try {
    context = await launchContext(profileDir, true, channel);

    const authenticated = await waitForOwaAuth(context, 5_000);
    if (!authenticated) {
      logger.debug('Headless: session invalid or expired (login redirect detected)');
      return null;
    }

    // Session is valid. Check if MSAL tokens are already there, or wait
    // up to 20s for OWA JS to silently acquire them (mirrors msteams-mcp).
    const tokensReady = await waitForMsalTokens(context, 20_000);
    if (!tokensReady) {
      logger.debug('Headless: MSAL tokens did not appear — falling back to headed login');
      return null;
    }

    const upn = await extractAndCacheTokens(context);
    if (!upn) return null;
    logger.info(`Headless login succeeded (${upn})`);
    return { upn, method: 'headless-sso' };
  } catch (err) {
    logger.debug('Headless login failed', err instanceof Error ? err.message : String(err));
    return null;
  } finally {
    await context?.close().catch(() => {});
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Headed login (fallback)
// ─────────────────────────────────────────────────────────────────────────────

/** Inject a visible progress overlay into the page (mirrors msteams-mcp). Failures are silently ignored — purely cosmetic. */
async function showOverlay(page: import('playwright').Page, phase: 'pending' | 'saving' | 'done' | 'error'): Promise<void> {
  const phases = {
    pending: { icon: '⋯', title: "You're signed in!", detail: 'Setting up your Outlook connection...',  bg: '#5b5fc7' },
    saving:  { icon: '⋯', title: 'Saving your session...',   detail: "So you won't need to log in again.", bg: '#5b5fc7' },
    done:    { icon: '✓', title: 'All done!',                 detail: 'This window will close automatically.', bg: '#107c10' },
    error:   { icon: '✕', title: 'Something went wrong',      detail: 'Please try again.',                 bg: '#c42b1c' },
  };
  const p = phases[phase];
  try {
    await page.evaluate(({ icon, title, detail, bg }) => {
      const existing = document.getElementById('msoutlook-mcp-overlay');
      if (existing) existing.remove();
      const overlay = document.createElement('div');
      overlay.id = 'msoutlook-mcp-overlay';
      Object.assign(overlay.style, { position: 'fixed', inset: '0', background: 'rgba(0,0,0,.65)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: '999999', fontFamily: "'Segoe UI',system-ui,sans-serif" });
      overlay.innerHTML = `<div style="background:white;border-radius:12px;padding:40px 48px;max-width:420px;text-align:center;box-shadow:0 8px 32px rgba(0,0,0,.3)"><div style="width:64px;height:64px;border-radius:50%;background:${bg};color:white;font-size:32px;display:flex;align-items:center;justify-content:center;margin:0 auto 24px">${icon}</div><h2 style="margin:0 0 12px;font-size:20px;font-weight:600;color:#242424">${title}</h2><p style="margin:0;font-size:14px;color:#616161;line-height:1.5">${detail}</p></div>`;
      document.body.appendChild(overlay);
    }, p);
  } catch { /* cosmetic — ignore */ }
}

export async function headedLogin(clearCookiesFirst = false): Promise<LoginResult | null> {
  logger.info('Opening browser for interactive login...');

  const profileDir = getBrowserProfileDir();
  const channel = getBrowserChannel();
  let context: BrowserContext | null = null;
  let browserClosed = false;

  try {
    context = await launchContext(profileDir, false, channel);
    context.on('close', () => { browserClosed = true; });

    // For force_new: clear browser cookies before importing fresh ones
    if (clearCookiesFirst) {
      await context.clearCookies();
      logger.debug('Browser cookies cleared for force_new login');
    }

    // Import Microsoft SSO cookies from the user's real browser — enables
    // instant silent sign-in without typing credentials (same as msteams-mcp).
    await importMicrosoftCookies(context, channel);

    const authenticated = await waitForOwaAuth(context, LOGIN_TIMEOUT_MS);
    const page = context.pages()[0];

    if (!authenticated) {
      // Login redirect — user needs to sign in manually; wait for them
      logger.info('Waiting for you to complete sign-in in the browser...');
      if (page) {
        await page.waitForFunction(owaAccessTokenPresent, { timeout: LOGIN_TIMEOUT_MS });
      }
    }

    // Session authenticated — show progress overlay while we save
    if (page) {
      await showOverlay(page, 'pending');
      await page.waitForTimeout(800);
      await showOverlay(page, 'saving');
    }

    const upn = await extractAndCacheTokens(context);
    if (!upn) {
      if (page) await showOverlay(page, 'error');
      return null;
    }

    if (page) {
      await showOverlay(page, 'done');
      await page.waitForTimeout(1200); // let user read "All done!"
    }

    logger.info(`Headed login succeeded (${upn}). Browser closing.`);
    return { upn, method: 'headed-browser' };
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
// Result type
// ─────────────────────────────────────────────────────────────────────────────

export interface LoginResult {
  upn: string;
  /** How authentication was completed. */
  method: 'token-cache' | 'headless-sso' | 'headed-browser';
}

// ─────────────────────────────────────────────────────────────────────────────
// Main entry point
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Login to Outlook Web.
 *
 * @param forceNew - When true, clear the saved session and force a full
 *   re-authentication (mirrors msteams-mcp's `forceNew` option).
 */
export async function browserLogin(forceNew = false): Promise<LoginResult | null> {
  if (forceNew) {
    clearSession();
    logger.info('Forced re-login — cleared previous session and token cache.');
  }

  if (!forceNew && hasSavedBrowserProfile()) {
    const result = await headlessLogin();
    if (result) return result;
    logger.info('Headless login failed — falling back to visible browser...');
  } else if (!forceNew) {
    logger.info('No saved browser profile — opening visible browser for first-time setup...');
  }

  // Pass clearCookiesFirst=true for force_new so the browser profile's cookies
  // are wiped before importing fresh SSO cookies (mirrors msteams-mcp forceNewLogin).
  return headedLogin(forceNew);
}

// ─────────────────────────────────────────────────────────────────────────────
// Token refresh alias
// ─────────────────────────────────────────────────────────────────────────────

export async function headlessTokenRefresh(): Promise<boolean> {
  return (await headlessLogin()) !== null;
}
