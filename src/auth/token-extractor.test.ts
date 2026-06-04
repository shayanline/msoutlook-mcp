import { describe, it, expect, vi, beforeEach } from 'vitest';
import {
  extractTokensFromLocalStorage,
  getOwaLocalStorage,
  type StorageState,
} from './token-extractor.js';
import { OWA_CLIENT_ID } from '../constants.js';

function b64url(obj: unknown): string {
  return Buffer.from(JSON.stringify(obj), 'utf8')
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
}

function makeJwt(payload: Record<string, unknown>): string {
  return `eyHEADER.${b64url(payload)}.sig`;
}

const futureSec = () => Math.floor(Date.now() / 1000) + 3600;
const farFutureSec = () => Math.floor(Date.now() / 1000) + 7200;
const pastSec = () => Math.floor(Date.now() / 1000) - 3600;

function entry(value: Record<string, unknown>) {
  return JSON.stringify(value);
}

describe('extractTokensFromLocalStorage', () => {
  beforeEach(() => {
    vi.spyOn(process.stderr, 'write').mockImplementation(() => true);
  });

  it('extracts owa, graph, refresh tokens plus upn and tenant', () => {
    const owaJwt = makeJwt({ exp: futureSec(), upn: 'me@x.com', tid: 'tenant-1' });
    const graphJwt = makeJwt({ exp: futureSec(), upn: 'me@x.com', tid: 'tenant-1' });
    const ls = [
      { name: 'unrelated', value: 'x' },
      {
        name: 'msal.3|acc|login.windows.net|accesstoken|cid|tenant|outlook',
        value: entry({ secret: owaJwt, credentialType: 'AccessToken', target: 'https://outlook.office.com/.default' }),
      },
      {
        name: 'msal.3|acc|login.windows.net|accesstoken|cid|tenant|graph',
        value: entry({ secret: graphJwt, credentialType: 'AccessToken', target: 'https://graph.microsoft.com/.default' }),
      },
      {
        name: 'msal.3|acc|login.windows.net|refreshtoken|cid|||',
        value: entry({ secret: 'REFRESH', credentialType: 'RefreshToken', clientId: OWA_CLIENT_ID }),
      },
    ];
    const result = extractTokensFromLocalStorage(ls);
    expect(result).not.toBeNull();
    expect(result!.owaToken).toBe(owaJwt);
    expect(result!.graphToken).toBe(graphJwt);
    expect(result!.refreshToken).toBe('REFRESH');
    expect(result!.upn).toBe('me@x.com');
    expect(result!.tenantId).toBe('tenant-1');
    expect(result!.owaTokenExpiry).toBeInstanceOf(Date);
    expect(result!.graphTokenExpiry).toBeInstanceOf(Date);
  });

  it('keeps the owa token with the latest expiry', () => {
    const older = makeJwt({ exp: futureSec() });
    const newer = makeJwt({ exp: farFutureSec() });
    const ls = [
      {
        name: 'msal.3|a|x|accesstoken|c|t|owa1',
        value: entry({ secret: older, target: 'https://outlook.office.com/.default' }),
      },
      {
        name: 'msal.3|a|x|accesstoken|c|t|owa2',
        value: entry({ secret: newer, target: 'https://outlook.office.com/.default' }),
      },
      {
        name: 'msal.3|a|x|refreshtoken|c|||',
        value: entry({ secret: 'R', clientId: OWA_CLIENT_ID }),
      },
    ];
    const result = extractTokensFromLocalStorage(ls);
    expect(result!.owaToken).toBe(newer);
  });

  it('falls back to preferred_username for upn', () => {
    const owaJwt = makeJwt({ exp: futureSec(), preferred_username: 'pref@x.com' });
    const ls = [
      {
        name: 'msal.3|a|x|accesstoken|c|t|owa',
        value: entry({ secret: owaJwt, target: 'https://outlook.office.com/.default' }),
      },
      {
        name: 'msal.3|a|x|refreshtoken|c|||',
        value: entry({ secret: 'R', clientId: OWA_CLIENT_ID }),
      },
    ];
    const result = extractTokensFromLocalStorage(ls);
    expect(result!.upn).toBe('pref@x.com');
    expect(result!.tenantId).toBeUndefined();
  });

  it('skips non-msal keys, bad JSON, missing secret, wrong-client refresh, non-jwt and expired tokens', () => {
    const owaJwt = makeJwt({ exp: futureSec() });
    const ls = [
      { name: 'notmsal', value: 'whatever' },
      { name: 'msal.bad', value: '{not json' },
      { name: 'msal.nosecret', value: entry({ credentialType: 'AccessToken' }) },
      { name: 'msal.3|a|x|refreshtoken|c|||', value: entry({ secret: 'R', clientId: 'OTHER' }) },
      { name: 'msal.3|a|x|accesstoken|c|t|owa', value: entry({ secret: 'plain-not-jwt', target: 'https://outlook.office.com/.default' }) },
      { name: 'msal.3|a|x|accesstoken|c|t|exp', value: entry({ secret: makeJwt({ exp: pastSec() }), target: 'https://outlook.office.com/.default' }) },
      { name: 'msal.3|a|x|accesstoken|c|t|noexp', value: entry({ secret: makeJwt({ foo: 'bar' }), target: 'https://outlook.office.com/.default' }) },
      { name: 'msal.3|a|x|accesstoken|c|t|noexpiry', value: entry({ secret: 'ey', target: 'https://outlook.office.com/.default' }) },
      { name: 'msal.3|a|x|accesstoken|c|t|badpayload', value: entry({ secret: 'ey.@@bad@@.sig', target: 'https://outlook.office.com/.default' }) },
      { name: 'msal.3|a|x|accesstoken|c|t|good', value: entry({ secret: owaJwt, target: 'https://outlook.office.com/.default' }) },
    ];
    const result = extractTokensFromLocalStorage(ls);
    // No valid refresh token (wrong clientId), so extraction fails overall
    expect(result).toBeNull();
  });

  it('returns null when there is no owa access token', () => {
    const ls = [
      {
        name: 'msal.3|a|x|refreshtoken|c|||',
        value: entry({ secret: 'R', clientId: OWA_CLIENT_ID }),
      },
    ];
    expect(extractTokensFromLocalStorage(ls)).toBeNull();
  });

  it('returns null when the refresh token is missing', () => {
    const owaJwt = makeJwt({ exp: futureSec() });
    const ls = [
      {
        name: 'msal.3|a|x|accesstoken|c|t|owa',
        value: entry({ secret: owaJwt, target: 'https://outlook.office.com/.default' }),
      },
    ];
    expect(extractTokensFromLocalStorage(ls)).toBeNull();
  });
});

describe('getOwaLocalStorage', () => {
  it('returns the localStorage array for the owa origin', () => {
    const state: StorageState = {
      origins: [
        { origin: 'https://other.com', localStorage: [{ name: 'a', value: '1' }] },
        { origin: 'https://outlook.office.com', localStorage: [{ name: 'b', value: '2' }] },
      ],
    };
    expect(getOwaLocalStorage(state)).toEqual([{ name: 'b', value: '2' }]);
  });

  it('returns null when no owa origin is present', () => {
    const state: StorageState = {
      origins: [{ origin: 'https://other.com', localStorage: [] }],
    };
    expect(getOwaLocalStorage(state)).toBeNull();
  });

  it('returns null when origins are undefined', () => {
    expect(getOwaLocalStorage({})).toBeNull();
  });
});
