import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('./client.js', () => ({ owaGet: vi.fn() }));

import { owaGet } from './client.js';
import { searchPeople } from './people.js';

const mGet = vi.mocked(owaGet);

beforeEach(() => vi.clearAllMocks());

describe('searchPeople', () => {
  it('searches with default top', async () => {
    mGet.mockResolvedValue({ value: [{ DisplayName: 'A' }] });
    const res = await searchPeople('jane');
    expect(res).toEqual([{ DisplayName: 'A' }]);
    const [path, params] = mGet.mock.calls[0] as [string, Record<string, string>];
    expect(path).toBe('/people');
    expect(params['$search']).toBe('"jane"');
    expect(params['$top']).toBe('10');
  });
  it('honours custom top', async () => {
    mGet.mockResolvedValue({ value: [] });
    await searchPeople('x', 3);
    expect((mGet.mock.calls[0][1] as Record<string, string>)['$top']).toBe('3');
  });
});
