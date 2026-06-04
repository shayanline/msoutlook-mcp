import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { TOKEN_REFRESH_BUFFER_MS } from '../constants.js';

const mocks = vi.hoisted(() => ({
  readTokenCache: vi.fn(),
  clearSession: vi.fn(),
  hasSessionState: vi.fn(),
  isSessionLikelyExpired: vi.fn(),
  refreshOwaToken: vi.fn(),
  refreshGraphToken: vi.fn(),
  headlessTokenRefresh: vi.fn(),
  browserLogin: vi.fn(),
}));

vi.mock('./session-store.js', () => ({
  readTokenCache: mocks.readTokenCache,
  clearSession: mocks.clearSession,
  hasSessionState: mocks.hasSessionState,
  isSessionLikelyExpired: mocks.isSessionLikelyExpired,
}));

vi.mock('./token-refresh.js', () => ({
  refreshOwaToken: mocks.refreshOwaToken,
  refreshGraphToken: mocks.refreshGraphToken,
}));

vi.mock('./browser-login.js', () => ({
  browserLogin: mocks.browserLogin,
  headlessTokenRefresh: mocks.headlessTokenRefresh,
}));

import {
  getOwaToken,
  getGraphToken,
  isAuthenticated,
  getAuthStatus,
} from './index.js';

const valid = (over: Record<string, unknown> = {}) => ({
  owaToken: 'owa',
  owaTokenExpiry: Date.now() + TOKEN_REFRESH_BUFFER_MS + 60_000,
  refreshToken: 'rt',
  tenantId: 'tid',
  upn: 'me@x.com',
  extractedAt: 1,
  ...over,
});

const expiring = (over: Record<string, unknown> = {}) =>
  valid({ owaTokenExpiry: Date.now() + 1000, ...over });

beforeEach(() => {
  vi.spyOn(process.stderr, 'write').mockImplementation(() => true);
  Object.values(mocks).forEach(m => m.mockReset());
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe('getOwaToken', () => {
  it('returns null when not authenticated', async () => {
    mocks.readTokenCache.mockReturnValue(null);
    expect(await getOwaToken()).toBeNull();
  });

  it('returns the cached token when still valid', async () => {
    mocks.readTokenCache.mockReturnValue(valid());
    expect(await getOwaToken()).toBe('owa');
    expect(mocks.refreshOwaToken).not.toHaveBeenCalled();
  });

  it('refreshes via HTTP when expiring', async () => {
    mocks.readTokenCache.mockReturnValue(expiring());
    mocks.refreshOwaToken.mockResolvedValue('fresh-owa');
    expect(await getOwaToken()).toBe('fresh-owa');
  });

  it('falls back to headless refresh and returns the new cached token', async () => {
    mocks.readTokenCache
      .mockReturnValueOnce(expiring())
      .mockReturnValueOnce(valid({ owaToken: 'browser-owa' }));
    mocks.refreshOwaToken.mockResolvedValue(null);
    mocks.headlessTokenRefresh.mockResolvedValue(true);
    expect(await getOwaToken()).toBe('browser-owa');
  });

  it('returns null when headless refresh succeeds but cache is empty', async () => {
    mocks.readTokenCache
      .mockReturnValueOnce(expiring())
      .mockReturnValueOnce(null);
    mocks.refreshOwaToken.mockResolvedValue(null);
    mocks.headlessTokenRefresh.mockResolvedValue(true);
    expect(await getOwaToken()).toBeNull();
  });

  it('returns null when all refresh methods fail', async () => {
    mocks.readTokenCache.mockReturnValue(expiring());
    mocks.refreshOwaToken.mockResolvedValue(null);
    mocks.headlessTokenRefresh.mockResolvedValue(false);
    expect(await getOwaToken()).toBeNull();
  });
});

describe('getGraphToken', () => {
  it('returns null when not authenticated', async () => {
    mocks.readTokenCache.mockReturnValue(null);
    expect(await getGraphToken()).toBeNull();
  });

  it('returns the cached graph token when valid', async () => {
    mocks.readTokenCache.mockReturnValue(valid({
      graphToken: 'g',
      graphTokenExpiry: Date.now() + TOKEN_REFRESH_BUFFER_MS + 60_000,
    }));
    expect(await getGraphToken()).toBe('g');
    expect(mocks.refreshGraphToken).not.toHaveBeenCalled();
  });

  it('refreshes when the graph token is missing', async () => {
    mocks.readTokenCache.mockReturnValue(valid());
    mocks.refreshGraphToken.mockResolvedValue('fresh-graph');
    expect(await getGraphToken()).toBe('fresh-graph');
  });

  it('refreshes when the graph token is expiring', async () => {
    mocks.readTokenCache.mockReturnValue(valid({
      graphToken: 'g',
      graphTokenExpiry: Date.now() + 1000,
    }));
    mocks.refreshGraphToken.mockResolvedValue('fresh-graph');
    expect(await getGraphToken()).toBe('fresh-graph');
  });

  it('returns null when the graph refresh fails', async () => {
    mocks.readTokenCache.mockReturnValue(valid());
    mocks.refreshGraphToken.mockResolvedValue(null);
    expect(await getGraphToken()).toBeNull();
  });
});

describe('isAuthenticated', () => {
  it('is false without a cache', () => {
    mocks.readTokenCache.mockReturnValue(null);
    expect(isAuthenticated()).toBe(false);
  });

  it('is true with a refresh token', () => {
    mocks.readTokenCache.mockReturnValue(valid());
    expect(isAuthenticated()).toBe(true);
  });

  it('is false without a refresh token', () => {
    mocks.readTokenCache.mockReturnValue(valid({ refreshToken: '' }));
    expect(isAuthenticated()).toBe(false);
  });
});

describe('getAuthStatus', () => {
  it('reports unauthenticated without a cache', () => {
    mocks.readTokenCache.mockReturnValue(null);
    expect(getAuthStatus()).toEqual({ authenticated: false });
  });

  it('reports full details including the graph token expiry', () => {
    const cache = valid({
      owaTokenExpiry: Date.now() + 30 * 60_000,
      graphTokenExpiry: Date.now() + 60 * 60_000,
    });
    mocks.readTokenCache.mockReturnValue(cache);
    const status = getAuthStatus();
    expect(status.authenticated).toBe(true);
    expect(status.upn).toBe('me@x.com');
    expect(status.tenantId).toBe('tid');
    expect(status.owaTokenMinutesRemaining).toBeGreaterThan(0);
    expect(status.graphTokenExpiry).toBeTypeOf('string');
  });

  it('clamps negative minutes to zero and omits the graph expiry when absent', () => {
    mocks.readTokenCache.mockReturnValue(valid({ owaTokenExpiry: Date.now() - 60_000 }));
    const status = getAuthStatus();
    expect(status.owaTokenMinutesRemaining).toBe(0);
    expect(status.graphTokenExpiry).toBeUndefined();
  });
});
