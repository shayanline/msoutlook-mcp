import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('../auth/index.js', () => ({
  getOwaToken: vi.fn(),
  getGraphToken: vi.fn(),
}));

vi.mock('../utils/http.js', () => ({
  getBearerHeaders: vi.fn(),
  parseResponse: vi.fn(),
  fetchWithRetry: vi.fn(),
}));

import { getOwaToken, getGraphToken } from '../auth/index.js';
import { getBearerHeaders, parseResponse, fetchWithRetry } from '../utils/http.js';
import {
  owaGet, owaPost, owaPatch, owaDelete,
  graphGet, graphPost, graphGetPath, graphGetBinary,
} from './client.js';
import { OWA_REST_V2, GRAPH_BASE, OWA_BASE } from '../constants.js';

const mGetOwaToken = vi.mocked(getOwaToken);
const mGetGraphToken = vi.mocked(getGraphToken);
const mGetBearerHeaders = vi.mocked(getBearerHeaders);
const mParseResponse = vi.mocked(parseResponse);
const mFetchWithRetry = vi.mocked(fetchWithRetry);

beforeEach(() => {
  vi.clearAllMocks();
  mGetBearerHeaders.mockReturnValue({ Authorization: 'Bearer x' } as Record<string, string>);
});

describe('owaGet', () => {
  it('throws when not authenticated', async () => {
    mGetOwaToken.mockResolvedValue(null as unknown as string);
    await expect(owaGet('/messages')).rejects.toThrow('Not authenticated. Run outlook_login first.');
    expect(mFetchWithRetry).not.toHaveBeenCalled();
  });

  it('fetches without params', async () => {
    mGetOwaToken.mockResolvedValue('tok');
    mFetchWithRetry.mockResolvedValue({} as Response);
    mParseResponse.mockResolvedValue({ ok: 1 });

    const result = await owaGet('/messages');

    expect(result).toEqual({ ok: 1 });
    expect(mGetBearerHeaders).toHaveBeenCalledWith('tok', OWA_BASE);
    expect(mFetchWithRetry).toHaveBeenCalledWith(`${OWA_REST_V2}/me/messages`, {
      method: 'GET',
      headers: { Authorization: 'Bearer x' },
    });
  });

  it('appends query string with ? when no existing query', async () => {
    mGetOwaToken.mockResolvedValue('tok');
    mFetchWithRetry.mockResolvedValue({} as Response);
    mParseResponse.mockResolvedValue([]);

    await owaGet('/messages', { '$top': '5' });

    const url = mFetchWithRetry.mock.calls[0][0] as string;
    expect(url).toBe(`${OWA_REST_V2}/me/messages?%24top=5`);
  });

  it('appends query string with & when path already has a query', async () => {
    mGetOwaToken.mockResolvedValue('tok');
    mFetchWithRetry.mockResolvedValue({} as Response);
    mParseResponse.mockResolvedValue([]);

    await owaGet('/messages?foo=bar', { a: 'b' });

    const url = mFetchWithRetry.mock.calls[0][0] as string;
    expect(url).toBe(`${OWA_REST_V2}/me/messages?foo=bar&a=b`);
  });
});

describe('owaPost', () => {
  it('throws when not authenticated', async () => {
    mGetOwaToken.mockResolvedValue(undefined as unknown as string);
    await expect(owaPost('/sendmail', {})).rejects.toThrow('Not authenticated. Run outlook_login first.');
  });

  it('posts JSON body', async () => {
    mGetOwaToken.mockResolvedValue('tok');
    mFetchWithRetry.mockResolvedValue({} as Response);
    mParseResponse.mockResolvedValue({ id: 'm1' });

    const result = await owaPost('/messages', { Subject: 'Hi' });

    expect(result).toEqual({ id: 'm1' });
    expect(mFetchWithRetry).toHaveBeenCalledWith(`${OWA_REST_V2}/me/messages`, {
      method: 'POST',
      headers: { Authorization: 'Bearer x' },
      body: JSON.stringify({ Subject: 'Hi' }),
    });
  });
});

describe('owaPatch', () => {
  it('throws when not authenticated', async () => {
    mGetOwaToken.mockResolvedValue(null as unknown as string);
    await expect(owaPatch('/messages/1', {})).rejects.toThrow('Not authenticated');
  });

  it('patches JSON body', async () => {
    mGetOwaToken.mockResolvedValue('tok');
    mFetchWithRetry.mockResolvedValue({} as Response);
    mParseResponse.mockResolvedValue({ ok: true });

    const result = await owaPatch('/messages/1', { IsRead: true });

    expect(result).toEqual({ ok: true });
    expect(mFetchWithRetry).toHaveBeenCalledWith(`${OWA_REST_V2}/me/messages/1`, {
      method: 'PATCH',
      headers: { Authorization: 'Bearer x' },
      body: JSON.stringify({ IsRead: true }),
    });
  });
});

describe('owaDelete', () => {
  it('throws when not authenticated', async () => {
    mGetOwaToken.mockResolvedValue(null as unknown as string);
    await expect(owaDelete('/messages/1')).rejects.toThrow('Not authenticated');
  });

  it('resolves on ok response', async () => {
    mGetOwaToken.mockResolvedValue('tok');
    mFetchWithRetry.mockResolvedValue({ ok: true, status: 200 } as Response);

    await expect(owaDelete('/messages/1')).resolves.toBeUndefined();
    expect(mFetchWithRetry).toHaveBeenCalledWith(`${OWA_REST_V2}/me/messages/1`, {
      method: 'DELETE',
      headers: { Authorization: 'Bearer x' },
    });
  });

  it('resolves on 204 even when not ok', async () => {
    mGetOwaToken.mockResolvedValue('tok');
    mFetchWithRetry.mockResolvedValue({ ok: false, status: 204 } as Response);

    await expect(owaDelete('/messages/1')).resolves.toBeUndefined();
  });

  it('throws on non-2xx with body text', async () => {
    mGetOwaToken.mockResolvedValue('tok');
    mFetchWithRetry.mockResolvedValue({
      ok: false,
      status: 404,
      text: vi.fn().mockResolvedValue('not found'),
    } as unknown as Response);

    await expect(owaDelete('/messages/1')).rejects.toThrow('HTTP 404: not found');
  });

  it('throws with empty text when text() rejects', async () => {
    mGetOwaToken.mockResolvedValue('tok');
    mFetchWithRetry.mockResolvedValue({
      ok: false,
      status: 500,
      text: vi.fn().mockRejectedValue(new Error('boom')),
    } as unknown as Response);

    await expect(owaDelete('/messages/1')).rejects.toThrow('HTTP 500: ');
  });
});

describe('graphGet', () => {
  it('throws when graph token unavailable', async () => {
    mGetGraphToken.mockResolvedValue(null as unknown as string);
    await expect(graphGet('/messages')).rejects.toThrow('Graph token unavailable. Run outlook_login first.');
  });

  it('fetches without params', async () => {
    mGetGraphToken.mockResolvedValue('gtok');
    mFetchWithRetry.mockResolvedValue({} as Response);
    mParseResponse.mockResolvedValue({ ok: 1 });

    await graphGet('/messages');

    expect(mGetBearerHeaders).toHaveBeenCalledWith('gtok');
    expect(mFetchWithRetry).toHaveBeenCalledWith(`${GRAPH_BASE}/me/messages`, {
      method: 'GET',
      headers: { Authorization: 'Bearer x' },
    });
  });

  it('appends params with ?', async () => {
    mGetGraphToken.mockResolvedValue('gtok');
    mFetchWithRetry.mockResolvedValue({} as Response);
    mParseResponse.mockResolvedValue({});

    await graphGet('/messages', { a: 'b' });

    expect(mFetchWithRetry.mock.calls[0][0]).toBe(`${GRAPH_BASE}/me/messages?a=b`);
  });

  it('appends params with & when query exists', async () => {
    mGetGraphToken.mockResolvedValue('gtok');
    mFetchWithRetry.mockResolvedValue({} as Response);
    mParseResponse.mockResolvedValue({});

    await graphGet('/messages?x=1', { a: 'b' });

    expect(mFetchWithRetry.mock.calls[0][0]).toBe(`${GRAPH_BASE}/me/messages?x=1&a=b`);
  });
});

describe('graphPost', () => {
  it('throws when graph token unavailable', async () => {
    mGetGraphToken.mockResolvedValue(null as unknown as string);
    await expect(graphPost('/x', {})).rejects.toThrow('Graph token unavailable');
  });

  it('posts JSON body', async () => {
    mGetGraphToken.mockResolvedValue('gtok');
    mFetchWithRetry.mockResolvedValue({} as Response);
    mParseResponse.mockResolvedValue({ ok: true });

    const result = await graphPost('/x', { a: 1 });

    expect(result).toEqual({ ok: true });
    expect(mFetchWithRetry).toHaveBeenCalledWith(`${GRAPH_BASE}/me/x`, {
      method: 'POST',
      headers: { Authorization: 'Bearer x' },
      body: JSON.stringify({ a: 1 }),
    });
  });
});

describe('graphGetPath', () => {
  it('throws when graph token unavailable', async () => {
    mGetGraphToken.mockResolvedValue(null as unknown as string);
    await expect(graphGetPath('/users/a@b.com')).rejects.toThrow('Graph token unavailable');
  });

  it('fetches arbitrary path without params', async () => {
    mGetGraphToken.mockResolvedValue('gtok');
    mFetchWithRetry.mockResolvedValue({} as Response);
    mParseResponse.mockResolvedValue({ id: 'u1' });

    await graphGetPath('/users/a@b.com');

    expect(mFetchWithRetry).toHaveBeenCalledWith(`${GRAPH_BASE}/users/a@b.com`, {
      method: 'GET',
      headers: { Authorization: 'Bearer x' },
    });
  });

  it('appends params with ?', async () => {
    mGetGraphToken.mockResolvedValue('gtok');
    mFetchWithRetry.mockResolvedValue({} as Response);
    mParseResponse.mockResolvedValue({});

    await graphGetPath('/me', { '$select': 'id' });

    expect(mFetchWithRetry.mock.calls[0][0]).toBe(`${GRAPH_BASE}/me?%24select=id`);
  });

  it('appends params with & when query exists', async () => {
    mGetGraphToken.mockResolvedValue('gtok');
    mFetchWithRetry.mockResolvedValue({} as Response);
    mParseResponse.mockResolvedValue({});

    await graphGetPath('/me?foo=1', { a: 'b' });

    expect(mFetchWithRetry.mock.calls[0][0]).toBe(`${GRAPH_BASE}/me?foo=1&a=b`);
  });
});

describe('graphGetBinary', () => {
  it('throws when graph token unavailable', async () => {
    mGetGraphToken.mockResolvedValue(null as unknown as string);
    await expect(graphGetBinary('/me/photo/$value')).rejects.toThrow('Graph token unavailable');
  });

  it('returns content type and bytes on ok', async () => {
    mGetGraphToken.mockResolvedValue('gtok');
    const buf = new Uint8Array([1, 2, 3]).buffer;
    mFetchWithRetry.mockResolvedValue({
      ok: true,
      headers: { get: vi.fn().mockReturnValue('image/png') },
      arrayBuffer: vi.fn().mockResolvedValue(buf),
    } as unknown as Response);

    const result = await graphGetBinary('/me/photo/$value');

    expect(result.contentType).toBe('image/png');
    expect(Buffer.isBuffer(result.bytes)).toBe(true);
    expect([...result.bytes]).toEqual([1, 2, 3]);
  });

  it('defaults content type when header missing', async () => {
    mGetGraphToken.mockResolvedValue('gtok');
    mFetchWithRetry.mockResolvedValue({
      ok: true,
      headers: { get: vi.fn().mockReturnValue(null) },
      arrayBuffer: vi.fn().mockResolvedValue(new Uint8Array([]).buffer),
    } as unknown as Response);

    const result = await graphGetBinary('/me/photo/$value');
    expect(result.contentType).toBe('application/octet-stream');
  });

  it('throws on non-ok with body text', async () => {
    mGetGraphToken.mockResolvedValue('gtok');
    mFetchWithRetry.mockResolvedValue({
      ok: false,
      status: 404,
      statusText: 'Not Found',
      text: vi.fn().mockResolvedValue('no photo'),
    } as unknown as Response);

    await expect(graphGetBinary('/me/photo/$value')).rejects.toThrow('HTTP 404 Not Found: no photo');
  });

  it('throws with empty text when text() rejects', async () => {
    mGetGraphToken.mockResolvedValue('gtok');
    mFetchWithRetry.mockResolvedValue({
      ok: false,
      status: 500,
      statusText: 'Server Error',
      text: vi.fn().mockRejectedValue(new Error('x')),
    } as unknown as Response);

    await expect(graphGetBinary('/me/photo/$value')).rejects.toThrow('HTTP 500 Server Error: ');
  });
});
