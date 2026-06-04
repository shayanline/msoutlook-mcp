import { describe, it, expect, vi, beforeEach } from 'vitest';
import * as directory from '../api/directory.js';
import * as mailbox from '../api/mailbox.js';
import { writeFile } from 'node:fs/promises';
import { registerDirectoryTools } from './directory-tools.js';

vi.mock('node:fs/promises', () => ({ writeFile: vi.fn() }));
vi.mock('../api/directory.js', () => ({ getUserProfile: vi.fn(), getManager: vi.fn(), getDirectReports: vi.fn(), getUserPhoto: vi.fn() }));
vi.mock('../api/mailbox.js', () => ({ getMailboxSettings: vi.fn(), setAutomaticReplies: vi.fn() }));

type Reg = { schema: any; handler: (a: any) => Promise<any> };
let tools: Map<string, Reg>;
function setup() {
  tools = new Map();
  const server = { tool: (n: string, _d: string, s: any, h: any) => tools.set(n, { schema: s, handler: h }) } as any;
  registerDirectoryTools(server);
}
const text = (r: any) => r.content[0].text;

beforeEach(() => { vi.clearAllMocks(); setup(); });

describe('get_user_profile', () => {
  it('full profile', async () => {
    vi.mocked(directory.getUserProfile).mockResolvedValue({
      displayName: 'Jane', mail: 'j@x.com', jobTitle: 'Eng', department: 'RD',
      officeLocation: 'HQ', mobilePhone: '111', businessPhones: ['222'], city: 'London', country: 'UK',
    } as any);
    const r = text(await tools.get('outlook_get_user_profile')!.handler({ email: 'j@x.com' }));
    expect(r).toContain('Name: Jane');
    expect(r).toContain('Office: HQ');
    expect(r).toContain('Mobile: 111');
    expect(r).toContain('Phone: 222');
    expect(r).toContain('Location: London, UK');
  });
  it('minimal profile (unknown name)', async () => {
    vi.mocked(directory.getUserProfile).mockResolvedValue({} as any);
    expect(text(await tools.get('outlook_get_user_profile')!.handler({}))).toBe('Name: (unknown)');
  });
});

describe('get_manager', () => {
  it('found', async () => {
    vi.mocked(directory.getManager).mockResolvedValue({ displayName: 'Boss' } as any);
    expect(text(await tools.get('outlook_get_manager')!.handler({ email: 'j@x.com' }))).toContain('Name: Boss');
  });
  it('throws -> no manager', async () => {
    vi.mocked(directory.getManager).mockRejectedValue(new Error('nope'));
    expect(text(await tools.get('outlook_get_manager')!.handler({}))).toContain('No manager found');
  });
});

describe('get_direct_reports', () => {
  it('reports', async () => {
    vi.mocked(directory.getDirectReports).mockResolvedValue([{ displayName: 'A' }, { displayName: 'B' }] as any);
    const r = text(await tools.get('outlook_get_direct_reports')!.handler({ email: 'j@x.com' }));
    expect(r).toContain('Report 1');
    expect(r).toContain('Report 2');
  });
  it('empty', async () => {
    vi.mocked(directory.getDirectReports).mockResolvedValue([] as any);
    expect(text(await tools.get('outlook_get_direct_reports')!.handler({}))).toBe('No direct reports.');
  });
});

describe('get_user_photo', () => {
  it('saves photo', async () => {
    vi.mocked(directory.getUserPhoto).mockResolvedValue({ bytes: Buffer.alloc(2048) } as any);
    const r = await tools.get('outlook_get_user_photo')!.handler({ output_path: '/p.jpg', email: 'j@x.com' });
    expect(writeFile).toHaveBeenCalledWith('/p.jpg', expect.anything());
    expect(text(r)).toContain('Saved photo');
  });
  it('throws -> no photo', async () => {
    vi.mocked(directory.getUserPhoto).mockRejectedValue(new Error('nope'));
    expect(text(await tools.get('outlook_get_user_photo')!.handler({ output_path: '/p.jpg' }))).toContain('No photo available');
  });
});

describe('get_automatic_replies', () => {
  it('scheduled with messages', async () => {
    vi.mocked(mailbox.getMailboxSettings).mockResolvedValue({
      AutomaticRepliesSetting: {
        Status: 'Scheduled',
        ScheduledStartDateTime: { DateTime: 's' }, ScheduledEndDateTime: { DateTime: 'e' },
        ExternalAudience: 'All', InternalReplyMessage: '<p>Internal</p>', ExternalReplyMessage: '<p>External</p>',
      },
    } as any);
    const r = text(await tools.get('outlook_get_automatic_replies')!.handler({}));
    expect(r).toContain('Automatic replies: Scheduled');
    expect(r).toContain('Window: s to e');
    expect(r).toContain('External audience: All');
    expect(r).toContain('Internal message: Internal');
    expect(r).toContain('External message: External');
  });
  it('disabled', async () => {
    vi.mocked(mailbox.getMailboxSettings).mockResolvedValue({ AutomaticRepliesSetting: { Status: 'Disabled' } } as any);
    expect(text(await tools.get('outlook_get_automatic_replies')!.handler({}))).toBe('Automatic replies: Disabled');
  });
  it('unknown when setting missing', async () => {
    vi.mocked(mailbox.getMailboxSettings).mockResolvedValue({} as any);
    expect(text(await tools.get('outlook_get_automatic_replies')!.handler({}))).toBe('Automatic replies: unknown.');
  });
});

describe('set_automatic_replies', () => {
  it('scheduled missing dates errors', async () => {
    expect(text(await tools.get('outlook_set_automatic_replies')!.handler({ status: 'Scheduled' }))).toContain('need both start and end');
  });
  it('always enabled success', async () => {
    vi.mocked(mailbox.setAutomaticReplies).mockResolvedValue({ AutomaticRepliesSetting: { Status: 'AlwaysEnabled', ExternalAudience: 'All' } } as any);
    const r = await tools.get('outlook_set_automatic_replies')!.handler({ status: 'AlwaysEnabled', internal_message: 'hi', external_message: 'ext', external_audience: 'All' });
    expect(text(r)).toContain('Automatic replies updated.');
    expect(mailbox.setAutomaticReplies).toHaveBeenCalled();
  });
  it('scheduled with dates success', async () => {
    vi.mocked(mailbox.setAutomaticReplies).mockResolvedValue({ AutomaticRepliesSetting: { Status: 'Scheduled', ScheduledStartDateTime: { DateTime: 's' }, ScheduledEndDateTime: { DateTime: 'e' } } } as any);
    const r = await tools.get('outlook_set_automatic_replies')!.handler({ status: 'Scheduled', start: 's', end: 'e', time_zone: 'UTC' });
    expect(text(r)).toContain('Window: s to e');
  });
});
