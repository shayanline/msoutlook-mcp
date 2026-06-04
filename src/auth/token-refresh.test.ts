import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

const store = vi.hoisted(() => ({
  readTokenCache: vi.fn(),
  writeTokenCache: vi.fn(),
}));

vi.mock('./session-store.js', () => ({
  readTokenCache: store.readTokenCache,
  writeTokenCache: store.writeTokenCache,
}));

import { refreshOwaToken, refreshGraphToken } from './token-refresh.js';

const baseCache = {
  owaToken: 'old-owa',
  owaTokenExpiry: 1000,
  refreshToken: 'rt',
  tenantId: 'tenant-1',
  extractedAt: 1,
};

function jsonResponse(body: unknown, ok = true, status = 200) {
  return {
    ok,
    status,
    json: async () => body,
    text: async () => JSON.stringify(body),
  } as unknown as Response;
}

beforeEach(() => {
  vi.spyOn(process.stderr, 'write').mockImplementation(() => true);
  store.readTokenCache.mockReset();
  store.writeTokenCache.mockReset();
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe('refreshOwaToken', () => {
  it('refreshes and caches a new token, rotating the refresh token', async () => {
    store.readTokenCache.mockReturnValue({ ...baseCache });
    const fetchMock = vi.fn().mockResolvedValue(
      jsonResponse({ access_token: 'new-owa', expires_in: 3600, refresh_token: 'new-rt' }),
    );
    vi.stubGlobal('fetch', fetchMock);

    const token = await refreshOwaToken();
    expect(token).toBe('new-owa');
    expect(store.writeTokenCache).toHaveBeenCalledTimes(1);
    const written = store.writeTokenCache.mock.calls[0][0];
    expect(written.owaToken).toBe('new-owa');
    expect(written.refreshToken).toBe('new-rt');

    // verify request shape
    const [url, init] = fetchMock.mock.calls[0];
    expect(url).toContain('/tenant-1/oauth2/v2.0/token');
    expect(init.method).toBe('POST');
    expect(init.headers.Origin).toBe('https://outlook.office.com');
  });

  it('keeps the existing refresh token when the response omits one', async () => {
    store.readTokenCache.mockReturnValue({ ...baseCache });
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(
      jsonResponse({ access_token: 'new-owa', expires_in: 3600 }),
    ));

    const token = await refreshOwaToken();
    expect(token).toBe('new-owa');
    expect(store.writeTokenCache.mock.calls[0][0].refreshToken).toBe('rt');
  });

  it('returns null when there is no cache', async () => {
    store.readTokenCache.mockReturnValue(null);
    expect(await refreshOwaToken()).toBeNull();
  });

  it('returns null when tenantId is missing', async () => {
    store.readTokenCache.mockReturnValue({ ...baseCache, tenantId: undefined });
    expect(await refreshOwaToken()).toBeNull();
  });

  it('returns null on a non-ok HTTP response', async () => {
    store.readTokenCache.mockReturnValue({ ...baseCache });
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(jsonResponse({ error: 'bad' }, false, 400)));
    expect(await refreshOwaToken()).toBeNull();
    expect(store.writeTokenCache).not.toHaveBeenCalled();
  });

  it('returns null on a non-ok response whose body cannot be read', async () => {
    store.readTokenCache.mockReturnValue({ ...baseCache });
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue({
      ok: false,
      status: 500,
      text: async () => { throw new Error('no body'); },
      json: async () => ({}),
    } as unknown as Response));
    expect(await refreshOwaToken()).toBeNull();
  });

  it('returns null and logs a timeout on AbortError', async () => {
    store.readTokenCache.mockReturnValue({ ...baseCache });
    const abort = Object.assign(new Error('aborted'), { name: 'AbortError' });
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(abort));
    expect(await refreshOwaToken()).toBeNull();
  });

  it('returns null on a generic network Error', async () => {
    store.readTokenCache.mockReturnValue({ ...baseCache });
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('netfail')));
    expect(await refreshOwaToken()).toBeNull();
  });

  it('returns null on a non-Error rejection', async () => {
    store.readTokenCache.mockReturnValue({ ...baseCache });
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue('weird'));
    expect(await refreshOwaToken()).toBeNull();
  });

  it('skips when a refresh is already in progress', async () => {
    store.readTokenCache.mockReturnValue({ ...baseCache });
    let release: (v: Response) => void = () => {};
    const hanging = new Promise<Response>(r => { release = r; });
    vi.stubGlobal('fetch', vi.fn().mockReturnValue(hanging));

    const first = refreshOwaToken();
    const second = await refreshOwaToken();
    expect(second).toBeNull();

    release(jsonResponse({ access_token: 'a', expires_in: 3600 }));
    expect(await first).toBe('a');
  });
});

describe('refreshGraphToken', () => {
  it('refreshes and caches the graph token', async () => {
    store.readTokenCache.mockReturnValue({ ...baseCache });
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(
      jsonResponse({ access_token: 'new-graph', expires_in: 3600, refresh_token: 'g-rt' }),
    ));
    const token = await refreshGraphToken();
    expect(token).toBe('new-graph');
    const written = store.writeTokenCache.mock.calls[0][0];
    expect(written.graphToken).toBe('new-graph');
    expect(written.graphTokenExpiry).toBeGreaterThan(Date.now());
  });

  it('returns null without a cached refresh token', async () => {
    store.readTokenCache.mockReturnValue({ ...baseCache, refreshToken: undefined });
    expect(await refreshGraphToken()).toBeNull();
  });

  it('returns null on a non-ok response', async () => {
    store.readTokenCache.mockReturnValue({ ...baseCache });
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(jsonResponse({}, false, 401)));
    expect(await refreshGraphToken()).toBeNull();
  });

  it('skips when a graph refresh is already in progress', async () => {
    store.readTokenCache.mockReturnValue({ ...baseCache });
    let release: (v: Response) => void = () => {};
    const hanging = new Promise<Response>(r => { release = r; });
    vi.stubGlobal('fetch', vi.fn().mockReturnValue(hanging));

    const first = refreshGraphToken();
    const second = await refreshGraphToken();
    expect(second).toBeNull();

    release(jsonResponse({ access_token: 'g', expires_in: 3600 }));
    expect(await first).toBe('g');
  });
});

describe('callTokenEndpoint timeout', () => {
  beforeEach(() => {
    vi.useFakeTimers();
    store.readTokenCache.mockReturnValue({ ...baseCache });
  });
  afterEach(() => {
    vi.useRealTimers();
    vi.unstubAllGlobals();
  });

  it('aborts the request once the timeout elapses', async () => {
    vi.stubGlobal('fetch', vi.fn().mockImplementation((_url: string, init: RequestInit) =>
      new Promise<Response>((_resolve, reject) => {
        (init.signal as AbortSignal).addEventListener('abort', () =>
          reject(Object.assign(new Error('aborted'), { name: 'AbortError' })));
      }),
    ));
    const pending = refreshOwaToken();
    await vi.advanceTimersByTimeAsync(10_000);
    expect(await pending).toBeNull();
  });
});
