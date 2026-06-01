/**
 * Import Microsoft SSO cookies from the user's real browser profile into a
 * Playwright context, enabling silent SSO without re-entering credentials.
 *
 * Ported and adapted from msteams-mcp's chrome-cookie-import module.
 * Supports both Chrome and Edge on macOS.
 *
 * Falls back gracefully if sqlite3 is missing, the browser isn't installed,
 * or the Keychain prompt is denied.
 */

import * as crypto from 'node:crypto';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { execSync } from 'node:child_process';
import type { BrowserContext } from 'playwright';
import { logger } from '../utils/logger.js';

// ─────────────────────────────────────────────────────────────────────────────
// Config
// ─────────────────────────────────────────────────────────────────────────────

const MICROSOFT_DOMAINS = [
  '%microsoftonline%',
  '%login.live.com%',
  '%login.microsoft.com%',
  '%microsoft.com%',
  '%office.com%',
  '%office365.com%',
  '%outlook.com%',
];

interface BrowserConfig {
  label: string;
  dataDir: string;
  keychainService: string;
  profileEnvVar: string;
}

const BROWSERS: Record<string, BrowserConfig> = {
  chrome: {
    label: 'Chrome',
    dataDir: path.join(os.homedir(), 'Library', 'Application Support', 'Google', 'Chrome'),
    keychainService: 'Chrome Safe Storage',
    profileEnvVar: 'MSOUTLOOK_CHROME_PROFILE',
  },
  msedge: {
    label: 'Edge',
    dataDir: path.join(os.homedir(), 'Library', 'Application Support', 'Microsoft Edge'),
    keychainService: 'Microsoft Edge Safe Storage',
    profileEnvVar: 'MSOUTLOOK_EDGE_PROFILE',
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Profile detection
// ─────────────────────────────────────────────────────────────────────────────

function listProfiles(dataDir: string): Array<{ dirName: string; name: string }> {
  const localStatePath = path.join(dataDir, 'Local State');
  if (!fs.existsSync(localStatePath)) return [];
  try {
    const state = JSON.parse(fs.readFileSync(localStatePath, 'utf8'));
    const infoCache = state?.profile?.info_cache;
    if (!infoCache || typeof infoCache !== 'object') return [];
    return Object.entries(infoCache).map(([dirName, info]) => ({
      dirName,
      name: String((info as Record<string, unknown>).name ?? ''),
    }));
  } catch {
    return [];
  }
}

function selectProfile(
  dataDir: string,
  envVar: string,
): string | null {
  const profiles = listProfiles(dataDir);

  // Explicit env var override
  const override = process.env[envVar];
  if (override) {
    return profiles.find(p => p.dirName === override)?.dirName ?? null;
  }

  // Default profile always exists
  if (fs.existsSync(path.join(dataDir, 'Default', 'Cookies'))) {
    return 'Default';
  }

  // Fall back to first available profile
  return profiles[0]?.dirName ?? null;
}

// ─────────────────────────────────────────────────────────────────────────────
// Cookie decryption (macOS AES-128-CBC via Keychain)
// ─────────────────────────────────────────────────────────────────────────────

const keyCache = new Map<string, Buffer | null>();

function getDecryptionKey(keychainService: string): Buffer | null {
  if (keyCache.has(keychainService)) return keyCache.get(keychainService) ?? null;

  try {
    const password = execSync(
      `security find-generic-password -s "${keychainService}" -w`,
      { encoding: 'utf8', timeout: 5000 },
    ).trim();
    const key = crypto.pbkdf2Sync(password, 'saltysalt', 1003, 16, 'sha1');
    keyCache.set(keychainService, key);
    return key;
  } catch {
    keyCache.set(keychainService, null);
    return null;
  }
}

function decryptCookieValue(hexValue: string, key: Buffer): string | null {
  try {
    const encrypted = Buffer.from(hexValue, 'hex');
    // v10 prefix = AES-128-CBC encrypted
    if (encrypted.length < 4 || encrypted[0] !== 0x76 || encrypted[1] !== 0x31 || encrypted[2] !== 0x30) {
      return encrypted.toString('utf8');
    }
    const ciphertext = encrypted.subarray(3);
    const iv = Buffer.alloc(16, 0x20);
    const decipher = crypto.createDecipheriv('aes-128-cbc', key, iv);
    decipher.setAutoPadding(true);
    return Buffer.concat([decipher.update(ciphertext), decipher.final()]).toString('utf8');
  } catch {
    return null;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Cookie reading (sqlite3 CLI)
// ─────────────────────────────────────────────────────────────────────────────

interface RawCookie {
  host_key: string;
  name: string;
  encrypted_value_hex: string;
  path: string;
  expires_utc: number;
  is_secure: number;
  is_httponly: number;
  samesite: number;
}

function readCookiesFromDb(dbPath: string): RawCookie[] {
  if (!fs.existsSync(dbPath)) return [];

  const tmpDb = path.join(os.tmpdir(), `msoutlook-mcp-cookies-${Date.now()}.db`);
  try {
    fs.copyFileSync(dbPath, tmpDb);
    for (const ext of ['-wal', '-shm']) {
      const src = dbPath + ext;
      if (fs.existsSync(src)) fs.copyFileSync(src, tmpDb + ext);
    }

    const where = MICROSOFT_DOMAINS.map(d => `host_key LIKE '${d}'`).join(' OR ');
    const sql = `SELECT host_key, name, hex(encrypted_value) as ev, path, expires_utc, is_secure, is_httponly, samesite FROM cookies WHERE (${where}) AND expires_utc > 0`;
    const output = execSync(`sqlite3 -separator '|||' "${tmpDb}" "${sql}"`, {
      encoding: 'utf8',
      timeout: 10_000,
      maxBuffer: 1024 * 1024,
    });

    return output
      .trim()
      .split('\n')
      .filter(line => line.includes('|||'))
      .map(line => {
        const [host_key, name, encrypted_value_hex, cookiePath, expires_utc, is_secure, is_httponly, samesite] = line.split('|||');
        return {
          host_key, name, encrypted_value_hex, path: cookiePath,
          expires_utc: parseInt(expires_utc, 10),
          is_secure: parseInt(is_secure, 10),
          is_httponly: parseInt(is_httponly, 10),
          samesite: parseInt(samesite, 10),
        };
      });
  } catch {
    return [];
  } finally {
    for (const f of [tmpDb, `${tmpDb}-wal`, `${tmpDb}-shm`]) {
      try { fs.unlinkSync(f); } catch { /* ignore */ }
    }
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Cookie conversion + injection
// ─────────────────────────────────────────────────────────────────────────────

function chromeEpochToUnix(ts: number): number {
  return Math.floor(ts / 1_000_000) - 11644473600;
}

function sameSiteLabel(v: number): 'Strict' | 'Lax' | 'None' {
  if (v === 2) return 'Strict';
  if (v === 1) return 'Lax';
  return 'None';
}

/**
 * Import Microsoft SSO cookies from the user's real browser into the Playwright context.
 * Only runs on macOS. Fails gracefully on all errors.
 *
 * @param context - The Playwright browser context to inject cookies into
 * @param channel - The browser channel being used ('chrome' or 'msedge')
 */
export async function importMicrosoftCookies(
  context: BrowserContext,
  channel: string | undefined,
): Promise<void> {
  if (process.platform !== 'darwin') return;

  // Determine which browser config to use based on channel
  const browserKey = channel === 'msedge' ? 'msedge' : 'chrome';
  const config = BROWSERS[browserKey];
  if (!config) return;

  if (!fs.existsSync(config.dataDir)) {
    logger.debug(`${config.label} not installed, skipping cookie import`);
    return;
  }

  const profileDir = selectProfile(config.dataDir, config.profileEnvVar);
  if (!profileDir) {
    logger.debug(`No ${config.label} profile found, skipping cookie import`);
    return;
  }

  const key = getDecryptionKey(config.keychainService);
  if (!key) {
    logger.debug(`Could not get ${config.label} decryption key from Keychain, skipping`);
    return;
  }

  const dbPath = path.join(config.dataDir, profileDir, 'Cookies');
  const rawCookies = readCookiesFromDb(dbPath);
  if (rawCookies.length === 0) {
    logger.debug(`No Microsoft cookies found in ${config.label} profile`);
    return;
  }

  const playwrightCookies = rawCookies
    .map(c => {
      const value = decryptCookieValue(c.encrypted_value_hex, key);
      if (!value) return null;
      return {
        name: c.name,
        value,
        domain: c.host_key.startsWith('.') ? c.host_key : `.${c.host_key}`,
        path: c.path || '/',
        expires: chromeEpochToUnix(c.expires_utc),
        secure: c.is_secure === 1,
        httpOnly: c.is_httponly === 1,
        sameSite: sameSiteLabel(c.samesite),
      };
    })
    .filter((c): c is NonNullable<typeof c> => c !== null);

  if (playwrightCookies.length === 0) return;

  try {
    await context.addCookies(playwrightCookies);
    logger.info(`Imported ${playwrightCookies.length} Microsoft SSO cookies from ${config.label} — browser will sign in automatically`);
  } catch (err) {
    logger.debug('Cookie injection failed', err instanceof Error ? err.message : String(err));
  }
}
