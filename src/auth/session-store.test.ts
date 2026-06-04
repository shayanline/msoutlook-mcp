import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { writeFileSync, utimesSync, existsSync, readFileSync } from 'node:fs';

const osState = vi.hoisted(() => ({
  home: `/tmp/outlook-store-${Date.now()}-${Math.random().toString(36).slice(2)}`,
  throwHome: false,
  throwUserInfo: false,
}));

vi.mock('node:os', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:os')>();
  return {
    ...actual,
    homedir: () => {
      if (osState.throwHome) throw new Error('no home');
      return osState.home;
    },
    userInfo: ((...args: unknown[]) => {
      if (osState.throwUserInfo) throw new Error('no user');
      // @ts-expect-error pass-through
      return actual.userInfo(...args);
    }) as typeof actual.userInfo,
  };
});

import * as store from './session-store.js';

const sampleCache = {
  owaToken: 'owa',
  owaTokenExpiry: Date.now() + 100000,
  graphToken: 'graph',
  graphTokenExpiry: Date.now() + 100000,
  refreshToken: 'rt',
  tenantId: 'tid',
  upn: 'me@x.com',
  extractedAt: Date.now(),
};

beforeEach(() => {
  vi.spyOn(process.stderr, 'write').mockImplementation(() => true);
});

afterEach(() => {
  store.clearSession();
  vi.restoreAllMocks();
});

describe('paths', () => {
  it('builds paths under the config directory', () => {
    expect(store.getSessionStatePath()).toContain('.msoutlook-mcp-server');
    expect(store.getTokenCachePath()).toContain('token-cache.json');
    expect(store.getBrowserProfileDir()).toContain('browser-profile');
    expect(store.CONFIG_DIR).toBe(osState.home + '/.msoutlook-mcp-server');
  });
});

describe('token cache round trip', () => {
  it('returns null when the cache file is missing', () => {
    expect(store.readTokenCache()).toBeNull();
  });

  it('writes encrypted and reads back the same data', () => {
    store.writeTokenCache(sampleCache);
    const raw = readFileSync(store.getTokenCachePath(), 'utf8');
    expect(raw).toContain('"iv"');
    expect(raw).toContain('"version": 1');
    expect(raw).not.toContain('owa');
    expect(store.readTokenCache()).toEqual(sampleCache);
  });

  it('clearTokenCache removes the file', () => {
    store.writeTokenCache(sampleCache);
    store.clearTokenCache();
    expect(store.readTokenCache()).toBeNull();
  });

  it('clearTokenCache is a no-op when nothing is cached', () => {
    expect(() => store.clearTokenCache()).not.toThrow();
  });
});

describe('session state', () => {
  it('round trips arbitrary state and reports presence', () => {
    expect(store.hasSessionState()).toBe(false);
    const state = { cookies: [{ name: 'a' }], origins: [] };
    store.writeSessionState(state);
    expect(store.hasSessionState()).toBe(true);
    expect(store.readSessionState()).toEqual(state);
  });

  it('reports session age and expiry', () => {
    expect(store.getSessionAge()).toBeNull();
    expect(store.isSessionLikelyExpired()).toBe(true);

    store.writeSessionState({ x: 1 });
    expect(store.getSessionAge()).toBeLessThan(1);
    expect(store.isSessionLikelyExpired()).toBe(false);

    const old = Date.now() / 1000 - 13 * 3600;
    utimesSync(store.getSessionStatePath(), old, old);
    expect(store.getSessionAge()).toBeGreaterThan(12);
    expect(store.isSessionLikelyExpired()).toBe(true);
  });
});

describe('legacy and error handling', () => {
  it('migrates a plaintext file to encrypted on read', () => {
    store.writeSessionState({ seed: true }); // ensures config dir exists
    const p = store.getTokenCachePath();
    writeFileSync(p, JSON.stringify(sampleCache), 'utf8');

    const read = store.readTokenCache();
    expect(read).toEqual(sampleCache);
    const migrated = readFileSync(p, 'utf8');
    expect(migrated).toContain('"iv"');
  });

  it('returns null when the file contains invalid JSON', () => {
    store.writeSessionState({ seed: true });
    writeFileSync(store.getTokenCachePath(), '{not valid json', 'utf8');
    expect(store.readTokenCache()).toBeNull();
  });
});

describe('clearSession', () => {
  it('removes the whole config directory', () => {
    store.writeTokenCache(sampleCache);
    expect(existsSync(store.CONFIG_DIR)).toBe(true);
    store.clearSession();
    expect(existsSync(store.CONFIG_DIR)).toBe(false);
    expect(store.hasSessionState()).toBe(false);
  });

  it('is a no-op when there is no config directory', () => {
    store.clearSession();
    expect(() => store.clearSession()).not.toThrow();
  });
});

describe('getConfigDir platform branches', () => {
  const originalPlatform = process.platform;

  function setPlatform(p: string) {
    Object.defineProperty(process, 'platform', { value: p, configurable: true });
  }

  afterEach(() => {
    Object.defineProperty(process, 'platform', { value: originalPlatform, configurable: true });
    osState.throwHome = false;
    osState.throwUserInfo = false;
    vi.unstubAllEnvs();
    vi.resetModules();
  });

  it('uses APPDATA on win32 when set', async () => {
    setPlatform('win32');
    vi.stubEnv('APPDATA', '/tmp/appdata-test');
    vi.resetModules();
    const fresh = await import('./session-store.js');
    expect(fresh.getSessionStatePath()).toContain('appdata-test');
    expect(fresh.getSessionStatePath()).toContain('msoutlook-mcp-server');
  });

  it('falls back to home AppData on win32 without APPDATA', async () => {
    setPlatform('win32');
    vi.stubEnv('APPDATA', undefined as unknown as string);
    vi.resetModules();
    const fresh = await import('./session-store.js');
    expect(fresh.getSessionStatePath()).toContain('AppData');
    expect(fresh.getSessionStatePath()).toContain('Roaming');
  });

  it('falls back to cwd data dir when there is no home', async () => {
    setPlatform('linux');
    osState.throwHome = true;
    vi.resetModules();
    const fresh = await import('./session-store.js');
    expect(fresh.getSessionStatePath()).toContain('msoutlook-mcp-server-data');
  });

  it('derives a key from CONFIG_DIR when userInfo throws', async () => {
    setPlatform('linux');
    osState.throwUserInfo = true;
    osState.home = `/tmp/outlook-store-userinfo-${Date.now()}`;
    vi.resetModules();
    const fresh = await import('./session-store.js');
    fresh.writeTokenCache(sampleCache);
    expect(fresh.readTokenCache()).toEqual(sampleCache);
    fresh.clearSession();
  });
});

describe('getConfigDir win32 without home', () => {
  const originalPlatform = process.platform;

  afterEach(() => {
    Object.defineProperty(process, 'platform', { value: originalPlatform, configurable: true });
    osState.throwHome = false;
    vi.unstubAllEnvs();
    vi.resetModules();
  });

  it('uses the cwd data dir when win32 lacks APPDATA and home', async () => {
    Object.defineProperty(process, 'platform', { value: 'win32', configurable: true });
    vi.stubEnv('APPDATA', undefined as unknown as string);
    osState.throwHome = true;
    vi.resetModules();
    const fresh = await import('./session-store.js');
    expect(fresh.getSessionStatePath()).toContain('msoutlook-mcp-server-data');
  });
});
