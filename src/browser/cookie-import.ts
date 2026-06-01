/**
 * Import Microsoft SSO cookies from the user's real browser profile into a
 * Playwright context, enabling silent SSO on first login without re-entering
 * credentials.
 *
 * Supports all major platforms:
 *   macOS  — Chrome/Edge via macOS Keychain, AES-128-CBC
 *   Linux  — Chrome/Edge via libsecret / "peanuts" fallback, AES-128-CBC
 *   Windows — Chrome/Edge via DPAPI (PowerShell), AES-256-GCM
 *
 * SQLite access uses sql.js (pure JS/WASM) — no native build tools needed.
 *
 * Fails gracefully on every error: cookie import is a best-effort optimisation.
 * If it fails the user falls back to a normal headed browser login.
 *
 * NOTE — Chrome 127+ on Windows uses App-Bound Encryption (prefix "APPB")
 * which cannot be decrypted outside the Chrome process. Those cookies are
 * silently skipped. Edge on Windows is unaffected and remains fully supported.
 */

import * as crypto from 'node:crypto';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { execSync } from 'node:child_process';
import type { BrowserContext } from 'playwright';
import { logger } from '../utils/logger.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

interface RawCookie {
  host_key: string;
  name: string;
  encrypted_value: Buffer;
  path: string;
  expires_utc: number;
  is_secure: number;
  is_httponly: number;
  samesite: number;
}

interface BrowserConfig {
  label: string;
  /** Base data directory (contains "Local State" and profile sub-dirs). */
  dataDir: string;
  /** macOS Keychain service name. */
  keychainService?: string;
  /** Env var to pin a specific profile dir name (e.g. "Profile 1"). */
  profileEnvVar: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// Per-platform browser configs
// ─────────────────────────────────────────────────────────────────────────────

function getBrowserConfigs(): Record<string, BrowserConfig> {
  const home = os.homedir();

  if (process.platform === 'darwin') {
    return {
      chrome: {
        label: 'Chrome',
        dataDir: path.join(home, 'Library', 'Application Support', 'Google', 'Chrome'),
        keychainService: 'Chrome Safe Storage',
        profileEnvVar: 'MSOUTLOOK_CHROME_PROFILE',
      },
      msedge: {
        label: 'Edge',
        dataDir: path.join(home, 'Library', 'Application Support', 'Microsoft Edge'),
        keychainService: 'Microsoft Edge Safe Storage',
        profileEnvVar: 'MSOUTLOOK_EDGE_PROFILE',
      },
    };
  }

  if (process.platform === 'win32') {
    const localAppData = process.env.LOCALAPPDATA ?? path.join(home, 'AppData', 'Local');
    return {
      chrome: {
        label: 'Chrome',
        dataDir: path.join(localAppData, 'Google', 'Chrome', 'User Data'),
        profileEnvVar: 'MSOUTLOOK_CHROME_PROFILE',
      },
      msedge: {
        label: 'Edge',
        dataDir: path.join(localAppData, 'Microsoft', 'Edge', 'User Data'),
        profileEnvVar: 'MSOUTLOOK_EDGE_PROFILE',
      },
    };
  }

  // Linux
  return {
    chrome: {
      label: 'Chrome',
      dataDir: path.join(home, '.config', 'google-chrome'),
      keychainService: 'Chrome Safe Storage',
      profileEnvVar: 'MSOUTLOOK_CHROME_PROFILE',
    },
    msedge: {
      label: 'Edge',
      dataDir: path.join(home, '.config', 'microsoft-edge'),
      keychainService: 'Microsoft Edge Safe Storage',
      profileEnvVar: 'MSOUTLOOK_EDGE_PROFILE',
    },
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Profile detection
// ─────────────────────────────────────────────────────────────────────────────

function selectProfile(dataDir: string, envVar: string): string | null {
  const override = process.env[envVar];
  if (override) return override;

  // Chrome/Edge newer versions store cookies at Default/Network/Cookies
  // Try Default profile first (most common), fall back to scanning
  const defaultPath = path.join(dataDir, 'Default');
  if (fs.existsSync(path.join(defaultPath, 'Network', 'Cookies')) ||
      fs.existsSync(path.join(defaultPath, 'Cookies'))) {
    return 'Default';
  }

  // Scan Local State for other profiles
  const localStatePath = path.join(dataDir, 'Local State');
  if (!fs.existsSync(localStatePath)) return null;
  try {
    const state = JSON.parse(fs.readFileSync(localStatePath, 'utf8')) as {
      profile?: { info_cache?: Record<string, unknown> };
    };
    const cache = state.profile?.info_cache ?? {};
    return Object.keys(cache)[0] ?? null;
  } catch {
    return null;
  }
}

/** Returns the Cookies database path, checking both old and new locations. */
function getCookiesDbPath(dataDir: string, profile: string): string | null {
  const base = path.join(dataDir, profile);
  // Newer Chrome/Edge moved cookies into a Network sub-directory
  const newPath = path.join(base, 'Network', 'Cookies');
  if (fs.existsSync(newPath)) return newPath;
  const oldPath = path.join(base, 'Cookies');
  if (fs.existsSync(oldPath)) return oldPath;
  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
// SQLite reading (sql.js — pure JS/WASM, no native deps)
// ─────────────────────────────────────────────────────────────────────────────

const MICROSOFT_DOMAINS_SQL = [
  '%microsoftonline%',
  '%login.live.com%',
  '%login.microsoft.com%',
  '%microsoft.com%',
  '%office.com%',
  '%office365.com%',
  '%outlook.com%',
].map(d => `host_key LIKE '${d}'`).join(' OR ');

async function readCookiesFromDb(dbPath: string): Promise<RawCookie[]> {
  // Copy to temp to avoid locking conflicts with a running browser
  const tmpDb = path.join(os.tmpdir(), `msoutlook-mcp-cookies-${Date.now()}.db`);
  try {
    fs.copyFileSync(dbPath, tmpDb);
    for (const ext of ['-wal', '-shm']) {
      const src = dbPath + ext;
      if (fs.existsSync(src)) fs.copyFileSync(src, tmpDb + ext);
    }

    // Dynamic import so sql.js WASM loads lazily (only when cookie import runs)
    const initSqlJs = (await import('sql.js')).default;
    const SQL = await initSqlJs();
    const fileBuffer = fs.readFileSync(tmpDb);
    const db = new SQL.Database(fileBuffer);

    const result = db.exec(
      `SELECT host_key, name, encrypted_value, path, expires_utc,
              is_secure, is_httponly, samesite
       FROM cookies
       WHERE (${MICROSOFT_DOMAINS_SQL}) AND expires_utc > 0`,
    );
    db.close();

    if (!result[0]) return [];

    const { columns, values } = result[0];
    const idx = (col: string) => columns.indexOf(col);

    return (values as unknown[][]).map(row => ({
      host_key: row[idx('host_key')] as string,
      name: row[idx('name')] as string,
      encrypted_value: Buffer.from(row[idx('encrypted_value')] as Uint8Array),
      path: row[idx('path')] as string,
      expires_utc: row[idx('expires_utc')] as number,
      is_secure: row[idx('is_secure')] as number,
      is_httponly: row[idx('is_httponly')] as number,
      samesite: row[idx('samesite')] as number,
    }));
  } catch (err) {
    logger.debug('Failed to read cookies DB', err instanceof Error ? err.message : String(err));
    return [];
  } finally {
    for (const f of [tmpDb, `${tmpDb}-wal`, `${tmpDb}-shm`]) {
      try { fs.unlinkSync(f); } catch { /* ignore */ }
    }
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// macOS key retrieval (Keychain AES-128-CBC)
// ─────────────────────────────────────────────────────────────────────────────

const macKeyCache = new Map<string, Buffer | null>();

function getMacOSDecryptionKey(keychainService: string): Buffer | null {
  if (macKeyCache.has(keychainService)) return macKeyCache.get(keychainService) ?? null;
  try {
    const password = execSync(
      `security find-generic-password -s "${keychainService}" -w`,
      { encoding: 'utf8', timeout: 5000 },
    ).trim();
    const key = crypto.pbkdf2Sync(password, 'saltysalt', 1003, 16, 'sha1');
    macKeyCache.set(keychainService, key);
    return key;
  } catch {
    macKeyCache.set(keychainService, null);
    return null;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Linux key retrieval (libsecret → "peanuts" fallback, AES-128-CBC)
// ─────────────────────────────────────────────────────────────────────────────

const linuxKeyCache = new Map<string, Buffer>();

function getLinuxDecryptionKey(keychainService: string): Buffer {
  if (linuxKeyCache.has(keychainService)) return linuxKeyCache.get(keychainService)!;

  let password = 'peanuts'; // fallback used when no keyring is available

  // Try libsecret first (GNOME / systemd)
  const appName = keychainService.toLowerCase().includes('edge') ? 'microsoft-edge' : 'chrome';
  const secretToolAttempts = [
    `secret-tool lookup xdg:schema chrome_libsecret_os_crypt_password_v2 application ${appName}`,
    `secret-tool lookup xdg:schema chrome_libsecret_os_crypt_password_v1 application ${appName}`,
    `secret-tool lookup service "${keychainService}" account "${appName}"`,
  ];
  for (const cmd of secretToolAttempts) {
    try {
      const result = execSync(cmd, { encoding: 'utf8', timeout: 3000 }).trim();
      if (result) { password = result; break; }
    } catch { /* continue */ }
  }

  // Linux uses 1 PBKDF2 iteration (vs 1003 on macOS)
  const key = crypto.pbkdf2Sync(password, 'saltysalt', 1, 16, 'sha1');
  linuxKeyCache.set(keychainService, key);
  return key;
}

// ─────────────────────────────────────────────────────────────────────────────
// Windows key retrieval (DPAPI via PowerShell + AES-256-GCM)
// ─────────────────────────────────────────────────────────────────────────────

let windowsAesKeyCache: Map<string, Buffer | null> | null = null;

function getWindowsDecryptionKey(dataDir: string): Buffer | null {
  if (!windowsAesKeyCache) windowsAesKeyCache = new Map();
  if (windowsAesKeyCache.has(dataDir)) return windowsAesKeyCache.get(dataDir) ?? null;

  try {
    const localStatePath = path.join(dataDir, 'Local State');
    if (!fs.existsSync(localStatePath)) return null;

    const localState = JSON.parse(fs.readFileSync(localStatePath, 'utf8')) as {
      os_crypt?: { encrypted_key?: string };
    };
    const encryptedKeyB64 = localState.os_crypt?.encrypted_key;
    if (!encryptedKeyB64) return null;

    // The key is base64-encoded; first 5 bytes are the literal ASCII "DPAPI" prefix
    const encryptedKeyBytes = Buffer.from(encryptedKeyB64, 'base64');
    const dpapiBlob = encryptedKeyBytes.subarray(5); // strip "DPAPI"
    const dpapiB64 = dpapiBlob.toString('base64');

    // Use PowerShell (always available on Windows 10+) to call CryptUnprotectData
    const psScript = `Add-Type -AssemblyName System.Security; [System.Convert]::ToBase64String([System.Security.Cryptography.ProtectedData]::Unprotect([System.Convert]::FromBase64String('${dpapiB64}'), $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser))`;
    const decryptedB64 = execSync(
      `powershell -NoProfile -NonInteractive -Command "${psScript}"`,
      { encoding: 'utf8', timeout: 10_000 },
    ).trim();

    const key = Buffer.from(decryptedB64, 'base64');
    windowsAesKeyCache.set(dataDir, key);
    return key;
  } catch (err) {
    logger.debug('Windows DPAPI key retrieval failed', err instanceof Error ? err.message : String(err));
    windowsAesKeyCache.set(dataDir, null);
    return null;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Cookie decryption (platform-aware)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Strip the 32-byte SHA256(host_key) domain-hash prefix that Chrome/Edge v24+
 * (2024) prepend to the plaintext before encryption. Older cookies have no
 * prefix, so we only strip when the first 32 bytes actually match the hash.
 */
function stripDomainHashPrefix(plaintext: Buffer, hostKey: string): Buffer {
  if (plaintext.length < 32) return plaintext;
  const domainHash = crypto.createHash('sha256').update(hostKey).digest();
  return plaintext.subarray(0, 32).equals(domainHash) ? plaintext.subarray(32) : plaintext;
}

function decryptCookieValue(
  encryptedValue: Buffer,
  platform: NodeJS.Platform,
  key: Buffer,
  hostKey: string,
): string | null {
  if (encryptedValue.length < 4) return null;

  const prefix = encryptedValue.subarray(0, 3).toString('ascii');

  // App-Bound Encryption (Chrome 127+ on Windows) — cannot decrypt externally
  if (encryptedValue.subarray(0, 4).toString('ascii') === 'APPB') {
    return null;
  }

  if (prefix !== 'v10' && prefix !== 'v11') {
    // Unencrypted plain-text value (legacy)
    return encryptedValue.toString('utf8').replace(/\0/g, '');
  }

  try {
    let plaintext: Buffer;
    if (platform === 'win32') {
      // Windows: AES-256-GCM
      // Layout: v10 (3 bytes) | nonce (12 bytes) | ciphertext | tag (16 bytes)
      // Minimum valid length = 3 + 12 + 16 = 31 bytes
      if (encryptedValue.length < 31) return null;
      const nonce = encryptedValue.subarray(3, 15);
      const ciphertextWithTag = encryptedValue.subarray(15);
      const tag = ciphertextWithTag.subarray(-16);
      const ciphertext = ciphertextWithTag.subarray(0, -16);

      const decipher = crypto.createDecipheriv('aes-256-gcm', key, nonce);
      decipher.setAuthTag(tag);
      plaintext = Buffer.concat([decipher.update(ciphertext), decipher.final()]);
    } else {
      // macOS / Linux: AES-128-CBC
      // Layout: v10/v11 (3 bytes) | ciphertext (rest)
      const ciphertext = encryptedValue.subarray(3);
      const iv = Buffer.alloc(16, 0x20); // 16 space characters
      const decipher = crypto.createDecipheriv('aes-128-cbc', key, iv);
      decipher.setAutoPadding(true);
      plaintext = Buffer.concat([decipher.update(ciphertext), decipher.final()]);
    }

    // Strip the v24+ domain-hash prefix if present, then decode.
    return stripDomainHashPrefix(plaintext, hostKey).toString('utf8');
  } catch {
    return null;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Cookie conversion helpers
// ─────────────────────────────────────────────────────────────────────────────

function chromeEpochToUnix(ts: number): number {
  return Math.floor(ts / 1_000_000) - 11644473600;
}

function sameSiteLabel(v: number): 'Strict' | 'Lax' | 'None' {
  if (v === 2) return 'Strict';
  if (v === 1) return 'Lax';
  return 'None';
}

// ─────────────────────────────────────────────────────────────────────────────
// Main export
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Import Microsoft SSO cookies from the user's real browser into the Playwright
 * context. Enables instant silent authentication — OWA recognises the session
 * and skips the login page entirely.
 *
 * @param context - Playwright browser context to inject cookies into
 * @param channel - Browser channel in use: 'chrome', 'msedge', or undefined
 */
export async function importMicrosoftCookies(
  context: BrowserContext,
  channel: string | undefined,
): Promise<void> {
  // Opt-out: users who'd rather sign in once manually (and avoid the one-time
  // Keychain/keyring prompt) can set this. The persistent profile then
  // remembers the session for all future logins.
  if (process.env.MSOUTLOOK_SKIP_COOKIE_IMPORT === 'true') {
    logger.debug('Cookie import skipped (MSOUTLOOK_SKIP_COOKIE_IMPORT=true)');
    return;
  }

  const platform = process.platform;

  const configs = getBrowserConfigs();
  const browserKey = channel === 'msedge' ? 'msedge' : 'chrome';
  const config = configs[browserKey];
  if (!config) return;

  if (!fs.existsSync(config.dataDir)) {
    logger.debug(`${config.label} data dir not found — skipping cookie import`);
    return;
  }

  const profile = selectProfile(config.dataDir, config.profileEnvVar);
  if (!profile) {
    logger.debug(`No ${config.label} profile found — skipping cookie import`);
    return;
  }

  const dbPath = getCookiesDbPath(config.dataDir, profile);
  if (!dbPath) {
    logger.debug(`No ${config.label} cookies database found — skipping cookie import`);
    return;
  }

  // Retrieve platform-specific decryption key
  let decryptionKey: Buffer | null = null;
  if (platform === 'darwin') {
    decryptionKey = getMacOSDecryptionKey(config.keychainService!);
  } else if (platform === 'linux') {
    decryptionKey = getLinuxDecryptionKey(config.keychainService ?? config.label);
  } else if (platform === 'win32') {
    decryptionKey = getWindowsDecryptionKey(config.dataDir);
  }

  if (!decryptionKey) {
    logger.debug(`Could not retrieve ${config.label} decryption key — skipping cookie import`);
    return;
  }

  const rawCookies = await readCookiesFromDb(dbPath);
  if (rawCookies.length === 0) {
    logger.debug(`No Microsoft cookies in ${config.label} profile`);
    return;
  }

  const nowSec = Math.floor(Date.now() / 1000);

  const playwrightCookies = rawCookies
    .map(c => {
      const value = decryptCookieValue(c.encrypted_value, platform, decryptionKey!, c.host_key);
      if (!value) return null;

      const expires = chromeEpochToUnix(c.expires_utc);
      if (expires <= nowSec) return null; // skip already-expired cookies

      const secure = c.is_secure === 1;
      let sameSite = sameSiteLabel(c.samesite);
      // Chromium rejects SameSite=None cookies that are not Secure ("Invalid cookie fields").
      // Downgrade those to Lax so the batch is always accepted.
      if (sameSite === 'None' && !secure) sameSite = 'Lax';

      return {
        name: c.name,
        value,
        domain: c.host_key, // raw host_key, as Chrome stores it (matches msteams-mcp)
        path: c.path || '/',
        expires,
        secure,
        httpOnly: c.is_httponly === 1,
        sameSite,
      };
    })
    .filter((c): c is NonNullable<typeof c> => c !== null);

  if (playwrightCookies.length === 0) {
    logger.debug(`No valid ${config.label} cookies to import (decryption failed or all expired)`);
    return;
  }

  // Try the whole batch first (fast path). If Playwright rejects it because a
  // single cookie has invalid fields, fall back to adding them one at a time so
  // one bad cookie can't block the rest.
  try {
    await context.addCookies(playwrightCookies);
    logger.info(`Imported ${playwrightCookies.length} Microsoft SSO cookies from ${config.label} (${platform}) — browser will sign in automatically`);
    return;
  } catch {
    logger.debug('Batch cookie inject failed — retrying individually');
  }

  let imported = 0;
  for (const cookie of playwrightCookies) {
    try {
      await context.addCookies([cookie]);
      imported++;
    } catch { /* skip the offending cookie */ }
  }

  if (imported > 0) {
    logger.info(`Imported ${imported}/${playwrightCookies.length} Microsoft SSO cookies from ${config.label} (${platform})`);
  } else {
    logger.debug('Cookie injection failed for all cookies');
  }
}
