import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('./client.js', () => ({
  owaGet: vi.fn(),
  owaPost: vi.fn(),
  owaDelete: vi.fn(),
}));

import { owaGet, owaPost, owaDelete } from './client.js';
import { listContacts, getContact, createContact, deleteContact } from './contacts.js';

const mGet = vi.mocked(owaGet);
const mPost = vi.mocked(owaPost);
const mDelete = vi.mocked(owaDelete);

beforeEach(() => vi.clearAllMocks());

describe('listContacts', () => {
  it('orders by name when not searching, default top', async () => {
    mGet.mockResolvedValue({ value: [{ Id: 'c1' }] });
    const res = await listContacts();
    expect(res).toEqual([{ Id: 'c1' }]);
    const [path, params] = mGet.mock.calls[0] as [string, Record<string, string>];
    expect(path).toBe('/contacts');
    expect(params['$top']).toBe('25');
    expect(params['$orderby']).toBe('displayName asc');
    expect(params['$search']).toBeUndefined();
  });
  it('uses $search and custom top when searching', async () => {
    mGet.mockResolvedValue({ value: [] });
    await listContacts('Jane', 5);
    const params = mGet.mock.calls[0][1] as Record<string, string>;
    expect(params['$search']).toBe('"Jane"');
    expect(params['$top']).toBe('5');
    expect(params['$orderby']).toBeUndefined();
  });
});

describe('getContact', () => {
  it('fetches by id', async () => {
    mGet.mockResolvedValue({ Id: 'c1' });
    await getContact('c1');
    expect(mGet).toHaveBeenCalledWith('/contacts/c1');
  });
});

describe('createContact', () => {
  it('creates with all fields', async () => {
    mPost.mockResolvedValue({ Id: 'c1' });
    await createContact({
      givenName: 'Jane', surname: 'Doe', email: 'j@d.com',
      businessPhone: '111', mobilePhone: '222', jobTitle: 'Eng', companyName: 'Acme', notes: 'vip',
    });
    const body = mPost.mock.calls[0][1] as any;
    expect(mPost.mock.calls[0][0]).toBe('/contacts');
    expect(body.DisplayName).toBe('Jane Doe');
    expect(body.EmailAddresses).toEqual([{ Address: 'j@d.com' }]);
    expect(body.BusinessPhones).toEqual(['111']);
    expect(body.MobilePhone1).toBe('222');
    expect(body.PersonalNotes).toBe('vip');
  });
  it('derives display name and omits email/phone when missing', async () => {
    mPost.mockResolvedValue({ Id: 'c2' });
    await createContact({ givenName: 'Solo' });
    const body = mPost.mock.calls[0][1] as any;
    expect(body.DisplayName).toBe('Solo');
    expect(body.EmailAddresses).toBeUndefined();
    expect(body.BusinessPhones).toBeUndefined();
  });
  it('uses explicit display name', async () => {
    mPost.mockResolvedValue({ Id: 'c3' });
    await createContact({ displayName: 'Custom Name' });
    expect((mPost.mock.calls[0][1] as any).DisplayName).toBe('Custom Name');
  });
});

describe('deleteContact', () => {
  it('deletes', async () => {
    mDelete.mockResolvedValue(undefined);
    await deleteContact('c1');
    expect(mDelete).toHaveBeenCalledWith('/contacts/c1');
  });
});

describe('createContact display name from surname only', () => {
  it('handles missing givenName', async () => {
    vi.mocked(owaPost).mockResolvedValue({ Id: 'c4' });
    await createContact({ surname: 'OnlyLast' });
    expect((vi.mocked(owaPost).mock.calls[0][1] as any).DisplayName).toBe('OnlyLast');
  });
});
