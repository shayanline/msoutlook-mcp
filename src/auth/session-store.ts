/**
 * Secure session state storage.
 *
 * Stores Playwright session state (localStorage, cookies) and cached tokens
 * in ~/.msoutlook-mcp-server/ (macOS/Linux) or %APPDATA%\msoutlook-mcp-server\ (Windows).
 *
 * Mirrors the storage approach from msteams-mcp:
 * - scryptSync key derivation (hostname:username — machine-specific, memory-hard)
 * - AES-256-GCM encryption at rest
 * - JSON envelope {iv, content, tag, version} for future-proof migration
 * - isEncrypted() check auto-migrates any legacy plaintext files
 */

import { scryptSync, createCipheriv, createDecipheriv, randomBytes } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, writeFileSync, rmSync, statSync } from 'node:fs';
import { homedir, hostname, userInfo } from 'node:os';
import { join } from 'node:path';
import { logger } from '../utils/logger.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

export interface TokenCache {
  owaToken: string;
  owaTokenExpiry: number;
  graphToken?: string;
  graphTokenExpiry?: number;
  refreshToken: string;
  tenantId?: string;
  upn?: string;
  extractedAt: number;
}

interface EncryptedEnvelope {
  iv: string;
  content: string;
  tag: string;
  version: 1;
}

// ─────────────────────────────────────────────────────────────────────────────
// Config directory (platform-aware, mirrors msteams-mcp)
// ─────────────────────────────────────────────────────────────────────────────

function getHomeDirSafe(): string | null {
  try { return homedir(); } catch { return null; }
}

function getConfigDir(): string {
  const home = getHomeDirSafe();

  if (process.platform === 'win32') {
    const appData = process.env.APPDATA ?? (home ? join(home, 'AppData', 'Roaming') : null);
    if (appData) return join(appData, 'msoutlook-mcp-server');
  }

  if (home) return join(home, '.msoutlook-mcp-server');

  // Fallback: alongside the dist directory
  return join(process.cwd(), 'msoutlook-mcp-server-data');
}

export const CONFIG_DIR = getConfigDir();
const SESSION_STATE_FILE = 'session-state.json';
const TOKEN_CACHE_FILE   = 'token-cache.json';
const BROWSER_PROFILE    = 'browser-profile';

/** Session considered stale after this many hours (matches msteams-mcp). */
const SESSION_EXPIRY_HOURS = 12;

// ─────────────────────────────────────────────────────────────────────────────
// Paths
// ─────────────────────────────────────────────────────────────────────────────

export function getSessionStatePath(): string  { return join(CONFIG_DIR, SESSION_STATE_FILE); }
export function getTokenCachePath(): string    { return join(CONFIG_DIR, TOKEN_CACHE_FILE); }
export function getBrowserProfileDir(): string { return join(CONFIG_DIR, BROWSER_PROFILE); }

function ensureConfigDir(): void {
  if (!existsSync(CONFIG_DIR)) mkdirSync(CONFIG_DIR, { recursive: true, mode: 0o700 });
}

// ─────────────────────────────────────────────────────────────────────────────
// Encryption — AES-256-GCM with scrypt-derived machine key (mirrors msteams-mcp)
// ─────────────────────────────────────────────────────────────────────────────

const SALT = 'msoutlook-mcp-credential-salt-v1';

function deriveKey(): Buffer {
  let machineId: string;
  try {
    machineId = `${hostname()}:${userInfo().username}`;
  } catch {
    machineId = CONFIG_DIR; // safe fallback
  }
  return scryptSync(machineId, SALT, 32) as Buffer;
}

function encryptJson(plaintext: string): EncryptedEnvelope {
  const key = deriveKey();
  const iv  = randomBytes(16);
  const cipher = createCipheriv('aes-256-gcm', key, iv);
  let content = cipher.update(plaintext, 'utf8', 'hex');
  content += cipher.final('hex');
  return { iv: iv.toString('hex'), content, tag: cipher.getAuthTag().toString('hex'), version: 1 };
}

function decryptEnvelope(env: EncryptedEnvelope): string {
  if (env.version !== 1) throw new Error(`Unsupported encryption version: ${env.version}`);
  const key = deriveKey();
  const iv  = Buffer.from(env.iv, 'hex');
  const tag = Buffer.from(env.tag, 'hex');
  const decipher = createDecipheriv('aes-256-gcm', key, iv);
  decipher.setAuthTag(tag);
  let out = decipher.update(env.content, 'hex', 'utf8');
  out += decipher.final('utf8');
  return out;
}

function isEncryptedEnvelope(v: unknown): v is EncryptedEnvelope {
  if (!v || typeof v !== 'object') return false;
  const o = v as Record<string, unknown>;
  return typeof o.iv === 'string' && typeof o.content === 'string'
      && typeof o.tag === 'string' && o.version === 1;
}

// ─────────────────────────────────────────────────────────────────────────────
// Generic read/write helpers
// ─────────────────────────────────────────────────────────────────────────────

function writeSecure(filePath: string, data: unknown): void {
  ensureConfigDir();
  const envelope = encryptJson(JSON.stringify(data));
  writeFileSync(filePath, JSON.stringify(envelope, null, 2), { mode: 0o600, encoding: 'utf8' });
}

function readSecure<T>(filePath: string): T | null {
  if (!existsSync(filePath)) return null;
  try {
    const parsed: unknown = JSON.parse(readFileSync(filePath, 'utf8'));
    if (isEncryptedEnvelope(parsed)) {
      return JSON.parse(decryptEnvelope(parsed)) as T;
    }
    // Legacy plaintext — migrate to encrypted in place
    logger.debug(`Migrating plaintext file to encrypted: ${filePath}`);
    writeSecure(filePath, parsed);
    return parsed as T;
  } catch (err) {
    logger.warn(`Failed to read ${filePath}`, err instanceof Error ? err.message : String(err));
    return null;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Session State (Playwright storageState)
// ─────────────────────────────────────────────────────────────────────────────

export function writeSessionState(state: unknown): void { writeSecure(getSessionStatePath(), state); }
export function readSessionState(): unknown | null      { return readSecure(getSessionStatePath()); }
export function hasSessionState(): boolean              { return existsSync(getSessionStatePath()); }

/** Returns the age of the session state file in hours, or null if it doesn't exist. */
export function getSessionAge(): number | null {
  const p = getSessionStatePath();
  if (!existsSync(p)) return null;
  return (Date.now() - statSync(p).mtimeMs) / (1000 * 60 * 60);
}

/** Returns true if the session state is missing or older than SESSION_EXPIRY_HOURS. */
export function isSessionLikelyExpired(): boolean {
  const age = getSessionAge();
  return age === null || age > SESSION_EXPIRY_HOURS;
}

// ─────────────────────────────────────────────────────────────────────────────
// Token Cache
// ─────────────────────────────────────────────────────────────────────────────

export function writeTokenCache(cache: TokenCache): void { writeSecure(getTokenCachePath(), cache); }
export function readTokenCache(): TokenCache | null      { return readSecure<TokenCache>(getTokenCachePath()); }
export function clearTokenCache(): void {
  const p = getTokenCachePath();
  if (existsSync(p)) { try { rmSync(p); } catch { /* ignore */ } }
}

// ─────────────────────────────────────────────────────────────────────────────
// Full session clear
// ─────────────────────────────────────────────────────────────────────────────

export function clearSession(): void {
  if (existsSync(CONFIG_DIR)) rmSync(CONFIG_DIR, { recursive: true, force: true });
}
