/**
 * HTTP utility helpers.
 */

import { OWA_BASE, OWA_USER_AGENT } from '../constants.js';

export function getBearerHeaders(token: string, origin = OWA_BASE): Record<string, string> {
  return {
    Authorization: `Bearer ${token}`,
    'Content-Type': 'application/json',
    Accept: 'application/json',
    Origin: origin,
    'User-Agent': OWA_USER_AGENT,
  };
}

/**
 * Parse a fetch Response, throwing on non-2xx.
 * JSON bodies are parsed; empty/non-JSON bodies (e.g. 202/204 from sendmail)
 * resolve to undefined.
 */
export async function parseResponse<T>(res: Response): Promise<T> {
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`HTTP ${res.status} ${res.statusText}: ${text.slice(0, 300)}`);
  }
  const ct = res.headers.get('content-type') ?? '';
  if (ct.includes('application/json')) {
    return res.json() as Promise<T>;
  }
  return undefined as T;
}

/** Sleep for ms milliseconds. */
export const sleep = (ms: number) => new Promise<void>(r => setTimeout(r, ms));

/** Retry a fetch operation with exponential backoff on 429/503/504. */
export async function fetchWithRetry(
  url: string,
  init: RequestInit,
  maxAttempts = 3,
): Promise<Response> {
  let lastError: Error | null = null;
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    const res = await fetch(url, init);
    if (res.status === 429 || res.status === 503 || res.status === 504) {
      const retryAfter = parseInt(res.headers.get('Retry-After') ?? '1', 10);
      const delay = (retryAfter || Math.pow(2, attempt)) * 1000 + Math.random() * 500;
      await sleep(delay);
      lastError = new Error(`HTTP ${res.status}`);
      continue;
    }
    return res;
  }
  throw lastError ?? new Error('Max retries exceeded');
}
