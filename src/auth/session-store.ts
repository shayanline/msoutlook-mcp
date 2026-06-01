/**
 * Session storage for auth tokens.
 *
 * Stores Playwright session state (localStorage, cookies) and cached tokens
 * in ~/.msoutlook-mcp-server/. Data is encrypted at rest using AES-256-GCM
 * with a machine-derived key.
 */

import { createCipheriv, createDecipheriv, randomBytes, createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, writeFileSync, rmSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { logger } from '../utils/logger.js';
import {
  SESSION_DIR_NAME,
  SESSION_STATE_FILE,
  TOKEN_CACHE_FILE,
  BROWSER_PROFILE_DIR,
} from '../constants.js';

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

// ─────────────────────────────────────────────────────────────────────────────
// Paths
// ─────────────────────────────────────────────────────────────────────────────

function getSessionDir(): string {
  return join(homedir(), SESSION_DIR_NAME);
}

export function getSessionStatePath(): string {
  return join(getSessionDir(), SESSION_STATE_FILE);
}

export function getTokenCachePath(): string {
  return join(getSessionDir(), TOKEN_CACHE_FILE);
}

export function getBrowserProfileDir(): string {
  return join(getSessionDir(), BROWSER_PROFILE_DIR);
}

function ensureSessionDir(): void {
  const dir = getSessionDir();
  if (!existsSync(dir)) {
    mkdirSync(dir, { recursive: true, mode: 0o700 });
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Encryption (AES-256-GCM with machine-derived key)
// ─────────────────────────────────────────────────────────────────────────────

function getDerivedKey(): Buffer {
  // Derive a key from a combination of the session dir path and hostname
  // Not perfect but prevents casual file reading; matches msteams-mcp approach
  const material = `msoutlook-mcp-${homedir()}-${process.platform}`;
  return createHash('sha256').update(material).digest();
}

function encrypt(plaintext: string): string {
  const key = getDerivedKey();
  const iv = randomBytes(12);
  const cipher = createCipheriv('aes-256-gcm', key, iv);
  const encrypted = Buffer.concat([cipher.update(plaintext, 'utf8'), cipher.final()]);
  const tag = cipher.getAuthTag();
  return Buffer.concat([iv, tag, encrypted]).toString('base64');
}

function decrypt(ciphertext: string): string {
  const key = getDerivedKey();
  const data = Buffer.from(ciphertext, 'base64');
  const iv = data.subarray(0, 12);
  const tag = data.subarray(12, 28);
  const encrypted = data.subarray(28);
  const decipher = createDecipheriv('aes-256-gcm', key, iv);
  decipher.setAuthTag(tag);
  return decipher.update(encrypted) + decipher.final('utf8');
}

// ─────────────────────────────────────────────────────────────────────────────
// Session State (Playwright storageState)
// ─────────────────────────────────────────────────────────────────────────────

export function writeSessionState(state: unknown): void {
  ensureSessionDir();
  const json = JSON.stringify(state);
  writeFileSync(getSessionStatePath(), encrypt(json), { mode: 0o600 });
}

export function readSessionState(): unknown | null {
  const path = getSessionStatePath();
  if (!existsSync(path)) return null;
  try {
    const raw = readFileSync(path, 'utf8');
    return JSON.parse(decrypt(raw));
  } catch (err) {
    logger.warn('Failed to read session state', err);
    return null;
  }
}

export function hasSessionState(): boolean {
  return existsSync(getSessionStatePath());
}

// ─────────────────────────────────────────────────────────────────────────────
// Token Cache
// ─────────────────────────────────────────────────────────────────────────────

export function writeTokenCache(cache: TokenCache): void {
  ensureSessionDir();
  const json = JSON.stringify(cache);
  writeFileSync(getTokenCachePath(), encrypt(json), { mode: 0o600 });
}

export function readTokenCache(): TokenCache | null {
  const path = getTokenCachePath();
  if (!existsSync(path)) return null;
  try {
    const raw = readFileSync(path, 'utf8');
    return JSON.parse(decrypt(raw)) as TokenCache;
  } catch (err) {
    logger.warn('Failed to read token cache', err);
    return null;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Clear
// ─────────────────────────────────────────────────────────────────────────────

export function clearSession(): void {
  const dir = getSessionDir();
  if (existsSync(dir)) {
    rmSync(dir, { recursive: true, force: true });
  }
}
