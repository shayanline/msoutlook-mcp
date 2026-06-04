import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('./client.js', () => ({
  graphGetPath: vi.fn(),
  graphGetBinary: vi.fn(),
}));

import { graphGetPath, graphGetBinary } from './client.js';
import { getUserProfile, getManager, getDirectReports, getUserPhoto } from './directory.js';

const mPath = vi.mocked(graphGetPath);
const mBin = vi.mocked(graphGetBinary);

beforeEach(() => vi.clearAllMocks());

describe('getUserProfile', () => {
  it('targets /me when no email', async () => {
    mPath.mockResolvedValue({ id: 'u1' });
    await getUserProfile();
    expect(mPath.mock.calls[0][0]).toBe('/me');
    expect((mPath.mock.calls[0][1] as Record<string, string>)['$select']).toContain('displayName');
  });
  it('targets and encodes the user email', async () => {
    mPath.mockResolvedValue({ id: 'u2' });
    await getUserProfile('a b@c.com');
    expect(mPath.mock.calls[0][0]).toBe('/users/a%20b%40c.com');
  });
});

describe('getManager', () => {
  it('appends /manager', async () => {
    mPath.mockResolvedValue({ id: 'm1' });
    await getManager('a@b.com');
    expect(mPath.mock.calls[0][0]).toBe('/users/a%40b.com/manager');
  });
});

describe('getDirectReports', () => {
  it('returns the value array', async () => {
    mPath.mockResolvedValue({ value: [{ id: 'r1' }] });
    const res = await getDirectReports();
    expect(res).toEqual([{ id: 'r1' }]);
    expect(mPath.mock.calls[0][0]).toBe('/me/directReports');
  });
  it('returns [] when value missing', async () => {
    mPath.mockResolvedValue({});
    const res = await getDirectReports('a@b.com');
    expect(res).toEqual([]);
    expect(mPath.mock.calls[0][0]).toBe('/users/a%40b.com/directReports');
  });
});

describe('getUserPhoto', () => {
  it('fetches binary photo', async () => {
    mBin.mockResolvedValue({ contentType: 'image/png', bytes: Buffer.from('x') });
    const res = await getUserPhoto('a@b.com');
    expect(res.contentType).toBe('image/png');
    expect(mBin.mock.calls[0][0]).toBe('/users/a%40b.com/photo/$value');
  });
});
