import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { getBearerHeaders, parseResponse, sleep, fetchWithRetry } from './http.js';
import { OWA_BASE, OWA_USER_AGENT } from '../constants.js';

function makeResponse(opts: {
  ok?: boolean;
  status?: number;
  statusText?: string;
  contentType?: string | null;
  json?: unknown;
  text?: string;
  retryAfter?: string | null;
}): Response {
  const headers = new Map<string, string>();
  if (opts.contentType !== undefined && opts.contentType !== null) {
    headers.set('content-type', opts.contentType);
  }
  if (opts.retryAfter !== undefined && opts.retryAfter !== null) {
    headers.set('Retry-After', opts.retryAfter);
  }
  return {
    ok: opts.ok ?? true,
    status: opts.status ?? 200,
    statusText: opts.statusText ?? 'OK',
    headers: { get: (k: string) => headers.get(k) ?? null },
    json: async () => opts.json,
    text: async () => opts.text ?? '',
  } as unknown as Response;
}

describe('getBearerHeaders', () => {
  it('builds headers with the default origin', () => {
    const h = getBearerHeaders('abc');
    expect(h).toEqual({
      Authorization: 'Bearer abc',
      'Content-Type': 'application/json',
      Accept: 'application/json',
      Origin: OWA_BASE,
      'User-Agent': OWA_USER_AGENT,
    });
  });

  it('uses a custom origin when provided', () => {
    const h = getBearerHeaders('xyz', 'https://example.com');
    expect(h.Origin).toBe('https://example.com');
    expect(h.Authorization).toBe('Bearer xyz');
  });
});

describe('parseResponse', () => {
  it('parses JSON bodies on 2xx', async () => {
    const res = makeResponse({ ok: true, contentType: 'application/json', json: { a: 1 } });
    await expect(parseResponse<{ a: number }>(res)).resolves.toEqual({ a: 1 });
  });

  it('returns undefined for empty bodies (no content-type)', async () => {
    const res = makeResponse({ ok: true, contentType: null });
    await expect(parseResponse(res)).resolves.toBeUndefined();
  });

  it('returns undefined for non-json content types', async () => {
    const res = makeResponse({ ok: true, contentType: 'text/plain' });
    await expect(parseResponse(res)).resolves.toBeUndefined();
  });

  it('throws on non-2xx including the body text', async () => {
    const res = makeResponse({ ok: false, status: 500, statusText: 'Server Error', text: 'boom' });
    await expect(parseResponse(res)).rejects.toThrow('HTTP 500 Server Error: boom');
  });

  it('throws on non-2xx even when reading the body fails', async () => {
    const res = {
      ok: false,
      status: 403,
      statusText: 'Forbidden',
      headers: { get: () => null },
      text: async () => { throw new Error('no body'); },
    } as unknown as Response;
    await expect(parseResponse(res)).rejects.toThrow('HTTP 403 Forbidden: ');
  });
});

describe('sleep', () => {
  it('resolves after the timer fires', async () => {
    vi.useFakeTimers();
    let done = false;
    const p = sleep(1000).then(() => { done = true; });
    await vi.advanceTimersByTimeAsync(1000);
    await p;
    expect(done).toBe(true);
    vi.useRealTimers();
  });
});

describe('fetchWithRetry', () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });
  afterEach(() => {
    vi.useRealTimers();
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it('returns the response immediately on success', async () => {
    const ok = makeResponse({ ok: true, status: 200 });
    const fetchMock = vi.fn().mockResolvedValue(ok);
    vi.stubGlobal('fetch', fetchMock);

    const res = await fetchWithRetry('https://x', { method: 'GET' });
    expect(res).toBe(ok);
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it('retries on 429 then succeeds, honouring Retry-After', async () => {
    const retry = makeResponse({ ok: false, status: 429, retryAfter: '2' });
    const ok = makeResponse({ ok: true, status: 200 });
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(retry)
      .mockResolvedValueOnce(ok);
    vi.stubGlobal('fetch', fetchMock);

    const promise = fetchWithRetry('https://x', { method: 'GET' });
    await vi.runAllTimersAsync();
    const res = await promise;

    expect(res).toBe(ok);
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it('uses exponential backoff when Retry-After is zero or missing', async () => {
    const retry503 = makeResponse({ ok: false, status: 503, retryAfter: '0' });
    const retry504 = makeResponse({ ok: false, status: 504 });
    const ok = makeResponse({ ok: true, status: 200 });
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(retry503)
      .mockResolvedValueOnce(retry504)
      .mockResolvedValueOnce(ok);
    vi.stubGlobal('fetch', fetchMock);

    const promise = fetchWithRetry('https://x', { method: 'GET' });
    await vi.runAllTimersAsync();
    const res = await promise;

    expect(res).toBe(ok);
    expect(fetchMock).toHaveBeenCalledTimes(3);
  });

  it('throws after max retries are exceeded', async () => {
    const retry = makeResponse({ ok: false, status: 503, retryAfter: '1' });
    const fetchMock = vi.fn().mockResolvedValue(retry);
    vi.stubGlobal('fetch', fetchMock);

    const promise = fetchWithRetry('https://x', { method: 'GET' }, 2);
    const assertion = expect(promise).rejects.toThrow('HTTP 503');
    await vi.runAllTimersAsync();
    await assertion;
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });
});
