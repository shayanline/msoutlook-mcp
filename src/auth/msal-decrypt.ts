/**
 * Decrypt MSAL Browser v4+ encrypted localStorage cache entries.
 *
 * Starting with MSAL browser v4, tokens in localStorage are AES-GCM encrypted
 * (unless the user selected "Keep me signed in"). The base key lives in the
 * session cookie `msal.cache.encryption`. This mirrors @azure/msal-browser's
 * BrowserCrypto.decrypt / LocalStorage.decryptData.
 */

import { webcrypto } from 'node:crypto';
import { logger } from '../utils/logger.js';
import { OWA_CLIENT_ID } from '../constants.js';

const ENCRYPTION_COOKIE = 'msal.cache.encryption';

export interface PlaywrightCookie {
  name: string;
  value: string;
  domain?: string;
}

interface EncryptedMsalEntry {
  id: string;
  nonce: string;
  data: string;
  lastUpdatedAt?: string;
}

interface MsalEncryptionKey {
  id: string;
  key: CryptoKey;
}

function base64DecToArr(base64String: string): Uint8Array {
  let encoded = base64String.replace(/-/g, '+').replace(/_/g, '/');
  switch (encoded.length % 4) {
    case 0:
      break;
    case 2:
      encoded += '==';
      break;
    case 3:
      encoded += '=';
      break;
    default:
      throw new Error('invalid base64');
  }
  return new Uint8Array(Buffer.from(encoded, 'base64'));
}

export function isEncryptedMsalEntry(entry: unknown): entry is EncryptedMsalEntry {
  if (!entry || typeof entry !== 'object') return false;
  const e = entry as Record<string, unknown>;
  return typeof e.id === 'string'
    && typeof e.nonce === 'string'
    && typeof e.data === 'string'
    && !e.secret;
}

function getEncryptionContext(key: string, clientId = OWA_CLIENT_ID): string {
  return key.includes(clientId) ? clientId : '';
}

async function generateHKDF(baseKey: Uint8Array): Promise<CryptoKey> {
  return webcrypto.subtle.importKey('raw', baseKey, 'HKDF', false, ['deriveKey']);
}

async function deriveKey(baseKey: CryptoKey, nonce: Uint8Array, context: string): Promise<CryptoKey> {
  return webcrypto.subtle.deriveKey(
    {
      name: 'HKDF',
      salt: nonce,
      hash: 'SHA-256',
      info: new TextEncoder().encode(context),
    },
    baseKey,
    { name: 'AES-GCM', length: 256 },
    false,
    ['encrypt', 'decrypt'],
  );
}

async function decryptPayload(
  baseKey: CryptoKey,
  nonce: string,
  context: string,
  encryptedData: string,
): Promise<string> {
  const encodedData = base64DecToArr(encryptedData);
  const derivedKey = await deriveKey(baseKey, base64DecToArr(nonce), context);
  const decrypted = await webcrypto.subtle.decrypt(
    {
      name: 'AES-GCM',
      iv: new Uint8Array(12),
    },
    derivedKey,
    encodedData,
  );
  return new TextDecoder().decode(decrypted);
}

/**
 * Parse the msal.cache.encryption cookie into an HKDF CryptoKey + key id.
 */
export async function loadMsalEncryptionKey(
  cookies: PlaywrightCookie[] | undefined,
): Promise<MsalEncryptionKey | null> {
  const raw = cookies?.find(c => c.name === ENCRYPTION_COOKIE)?.value;
  if (!raw) return null;

  try {
    const parsed = JSON.parse(decodeURIComponent(raw)) as { key?: string; id?: string };
    if (!parsed?.key || !parsed?.id) return null;
    const baseKey = base64DecToArr(parsed.key);
    return {
      id: parsed.id,
      key: await generateHKDF(baseKey),
    };
  } catch (err) {
    logger.debug(
      'Failed to parse msal.cache.encryption cookie',
      err instanceof Error ? err.message : String(err),
    );
    return null;
  }
}

/**
 * Resolve MSAL localStorage entries, decrypting v4 encrypted payloads when possible.
 * Returns a new array of {name, value} suitable for extractTokensFromLocalStorage.
 */
export async function resolveMsalLocalStorage(
  localStorage: Array<{ name: string; value: string }>,
  cookies?: PlaywrightCookie[],
): Promise<Array<{ name: string; value: string }>> {
  const encryption = await loadMsalEncryptionKey(cookies);
  let decryptedCount = 0;
  let skippedExpired = 0;
  let decryptFailures = 0;
  const resolved: Array<{ name: string; value: string }> = [];

  for (const item of localStorage) {
    const key = item.name;
    if (!key.startsWith('msal.')) {
      resolved.push(item);
      continue;
    }

    let entry: unknown;
    try {
      entry = JSON.parse(item.value);
    } catch {
      resolved.push(item);
      continue;
    }

    if (!isEncryptedMsalEntry(entry)) {
      resolved.push(item);
      continue;
    }

    if (!encryption) {
      decryptFailures++;
      continue;
    }
    if (entry.id !== encryption.id) {
      skippedExpired++;
      continue;
    }

    try {
      const plaintext = await decryptPayload(
        encryption.key,
        entry.nonce,
        getEncryptionContext(key),
        entry.data,
      );
      const decrypted = JSON.parse(plaintext) as Record<string, unknown>;
      if (entry.lastUpdatedAt && !decrypted.lastUpdatedAt) {
        decrypted.lastUpdatedAt = entry.lastUpdatedAt;
      }
      resolved.push({ name: key, value: JSON.stringify(decrypted) });
      decryptedCount++;
    } catch (err) {
      decryptFailures++;
      logger.debug(
        'MSAL decrypt failed for key',
        key.slice(0, 80),
        err instanceof Error ? err.message : String(err),
      );
    }
  }

  if (decryptedCount || skippedExpired || decryptFailures) {
    logger.debug('MSAL localStorage decrypt summary', {
      decryptedCount,
      skippedExpired,
      decryptFailures,
      hasEncryptionCookie: !!encryption,
    });
  }

  return resolved;
}
