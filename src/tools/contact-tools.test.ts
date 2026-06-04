import { describe, it, expect, vi, beforeEach } from 'vitest';
import * as contacts from '../api/contacts.js';
import * as people from '../api/people.js';
import * as presence from '../api/presence.js';
import { registerContactTools } from './contact-tools.js';

vi.mock('../api/contacts.js', () => ({ listContacts: vi.fn(), getContact: vi.fn(), createContact: vi.fn(), deleteContact: vi.fn() }));
vi.mock('../api/people.js', () => ({ searchPeople: vi.fn() }));
vi.mock('../api/presence.js', () => ({ getAvailability: vi.fn() }));

type Reg = { schema: any; handler: (a: any) => Promise<any> };
let tools: Map<string, Reg>;
function setup() {
  tools = new Map();
  const server = { tool: (n: string, _d: string, s: any, h: any) => tools.set(n, { schema: s, handler: h }) } as any;
  registerContactTools(server);
}
const text = (r: any) => r.content[0].text;

beforeEach(() => { vi.clearAllMocks(); setup(); });

describe('list_contacts', () => {
  it('full and minimal contacts', async () => {
    vi.mocked(contacts.listContacts).mockResolvedValue([
      { Id: 'c1', DisplayName: 'Jane', EmailAddresses: [{ Address: 'j@x.com' }], BusinessPhones: ['111'], MobilePhone1: '222', JobTitle: 'Eng', CompanyName: 'Acme', Department: 'RD' },
      { Id: 'c2', DisplayName: 'Bob' },
    ] as any);
    const r = text(await tools.get('outlook_list_contacts')!.handler({ search: 'a', top: 5 }));
    expect(contacts.listContacts).toHaveBeenCalledWith('a', 5);
    expect(r).toContain('Email: j@x.com');
    expect(r).toContain('Phone: 111, 222');
    expect(r).toContain('Title: Eng');
    expect(r).toContain('Company: Acme');
    expect(r).toContain('Department: RD');
    expect(r).toContain('Name: Bob');
  });
  it('empty + default top', async () => {
    vi.mocked(contacts.listContacts).mockResolvedValue([] as any);
    const r = await tools.get('outlook_list_contacts')!.handler({});
    expect(contacts.listContacts).toHaveBeenCalledWith(undefined, 25);
    expect(text(r)).toBe('No contacts found.');
  });
});

describe('get/create/delete contact', () => {
  it('get_contact', async () => {
    vi.mocked(contacts.getContact).mockResolvedValue({ Id: 'c', DisplayName: 'Jane' } as any);
    expect(text(await tools.get('outlook_get_contact')!.handler({ id: 'c' }))).toContain('Name: Jane');
  });
  it('create_contact', async () => {
    vi.mocked(contacts.createContact).mockResolvedValue({ Id: 'c', DisplayName: 'Jane' } as any);
    const r = await tools.get('outlook_create_contact')!.handler({ given_name: 'Jane', surname: 'Doe', email: 'j@x.com', business_phone: '1', mobile_phone: '2', job_title: 'Eng', company_name: 'Acme', notes: 'n' });
    expect(text(r)).toContain('Contact created.');
    expect(contacts.createContact).toHaveBeenCalled();
  });
  it('delete_contact', async () => {
    vi.mocked(contacts.deleteContact).mockResolvedValue(undefined as any);
    expect(text(await tools.get('outlook_delete_contact')!.handler({ id: 'c' }))).toBe('Contact deleted.');
  });
});

describe('search_people', () => {
  it('found with scored, fallback, none email', async () => {
    vi.mocked(people.searchPeople).mockResolvedValue([
      { DisplayName: 'A', ScoredEmailAddresses: [{ Address: 'a@x.com' }], UserPrincipalName: 'a@x.com', JobTitle: 'Eng', Department: 'RD', CompanyName: 'Acme' },
      { DisplayName: 'B', EmailAddresses: [{ Address: 'b@x.com' }] },
      { DisplayName: 'C' },
    ] as any);
    const r = text(await tools.get('outlook_search_people')!.handler({ query: 'q', top: 5 }));
    expect(people.searchPeople).toHaveBeenCalledWith('q', 5);
    expect(r).toContain('Email: a@x.com');
    expect(r).toContain('UPN: a@x.com');
    expect(r).toContain('Email: b@x.com');
    expect(r).toContain('Name: C');
  });
  it('empty + default top', async () => {
    vi.mocked(people.searchPeople).mockResolvedValue([] as any);
    const r = await tools.get('outlook_search_people')!.handler({ query: 'q' });
    expect(people.searchPeople).toHaveBeenCalledWith('q', 10);
    expect(text(r)).toBe('No people found.');
  });
});

describe('get_availability', () => {
  it('covers all availability branches', async () => {
    vi.mocked(presence.getAvailability).mockResolvedValue([
      {
        displayName: 'Jane', email: 'j@x.com', status: 'Busy', statusChangesAt: '12:00',
        outOfOffice: { isActive: true, scheduledStart: 's', scheduledEnd: 'e', message: '<p>Away now</p>' },
        workingHours: { startTime: '09:00', endTime: '17:00', daysOfWeek: ['Mon', 'Tue'], timeZone: 'UTC' },
      },
      {
        email: 'b@x.com', status: undefined, scheduleError: 'err',
        outOfOffice: { isActive: false },
      },
      {
        email: 'c@x.com', status: 'Free',
        outOfOfficeError: 'ooferr',
        workingHours: { startTime: '09:00', endTime: '17:00' },
      },
    ] as any);
    const r = text(await tools.get('outlook_get_availability')!.handler({ emails: ['j@x.com'], window_hours: 4 }));
    expect(presence.getAvailability).toHaveBeenCalledWith(['j@x.com'], 4);
    expect(r).toContain('Jane <j@x.com>');
    expect(r).toContain('Free/busy now: Busy until 12:00');
    expect(r).toContain('Out of office: yes (s to e)');
    expect(r).toContain('Auto reply: Away now');
    expect(r).toContain('Working hours: 09:00 to 17:00 UTC (Mon, Tue)');
    expect(r).toContain('Free/busy now: unavailable (err)');
    expect(r).toContain('Out of office: no');
    expect(r).toContain('Free/busy now: Free');
    expect(r).toContain('Out of office: unknown (ooferr)');
  });
  it('default window hours', async () => {
    vi.mocked(presence.getAvailability).mockResolvedValue([{ email: 'a@x.com', status: 'Busy', outOfOffice: { isActive: true } }] as any);
    await tools.get('outlook_get_availability')!.handler({ emails: ['a@x.com'] });
    expect(presence.getAvailability).toHaveBeenCalledWith(['a@x.com'], 8);
  });
});
