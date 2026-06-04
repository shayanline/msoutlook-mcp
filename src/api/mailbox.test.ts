import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('./client.js', () => ({
  owaGet: vi.fn(),
  owaPatch: vi.fn(),
}));

import { owaGet, owaPatch } from './client.js';
import { getMailboxSettings, setAutomaticReplies } from './mailbox.js';

const mGet = vi.mocked(owaGet);
const mPatch = vi.mocked(owaPatch);

beforeEach(() => vi.clearAllMocks());

describe('getMailboxSettings', () => {
  it('reads settings', async () => {
    mGet.mockResolvedValue({ TimeZone: 'UTC' });
    const res = await getMailboxSettings();
    expect(res).toEqual({ TimeZone: 'UTC' });
    expect(mGet).toHaveBeenCalledWith('/mailboxsettings');
  });
});

describe('setAutomaticReplies', () => {
  it('disables without message/schedule and defaults audience to All', async () => {
    mPatch.mockResolvedValue({});
    await setAutomaticReplies({ status: 'Disabled' });
    const body = mPatch.mock.calls[0][1] as any;
    expect(mPatch.mock.calls[0][0]).toBe('/mailboxsettings');
    expect(body.AutomaticRepliesSetting.Status).toBe('Disabled');
    expect(body.AutomaticRepliesSetting.ExternalAudience).toBe('All');
    expect(body.AutomaticRepliesSetting.InternalReplyMessage).toBeUndefined();
    expect(body.AutomaticRepliesSetting.ScheduledStartDateTime).toBeUndefined();
  });

  it('enables always with internal message defaulting external to internal', async () => {
    mPatch.mockResolvedValue({});
    await setAutomaticReplies({ status: 'AlwaysEnabled', internalMessage: 'Away', externalAudience: 'ContactsOnly' });
    const s = (mPatch.mock.calls[0][1] as any).AutomaticRepliesSetting;
    expect(s.ExternalAudience).toBe('ContactsOnly');
    expect(s.InternalReplyMessage).toBe('Away');
    expect(s.ExternalReplyMessage).toBe('Away');
  });

  it('defaults internal/external to empty string when omitted', async () => {
    mPatch.mockResolvedValue({});
    await setAutomaticReplies({ status: 'AlwaysEnabled' });
    const s = (mPatch.mock.calls[0][1] as any).AutomaticRepliesSetting;
    expect(s.InternalReplyMessage).toBe('');
    expect(s.ExternalReplyMessage).toBe('');
  });

  it('uses distinct external message when supplied', async () => {
    mPatch.mockResolvedValue({});
    await setAutomaticReplies({ status: 'AlwaysEnabled', internalMessage: 'In', externalMessage: 'Ext' });
    const s = (mPatch.mock.calls[0][1] as any).AutomaticRepliesSetting;
    expect(s.ExternalReplyMessage).toBe('Ext');
  });

  it('scheduled with start and end and custom timezone', async () => {
    mPatch.mockResolvedValue({});
    await setAutomaticReplies({ status: 'Scheduled', internalMessage: 'OOO', start: '2024-01-01T09:00:00', end: '2024-01-05T17:00:00', timeZone: 'Europe/London' });
    const s = (mPatch.mock.calls[0][1] as any).AutomaticRepliesSetting;
    expect(s.ScheduledStartDateTime).toEqual({ DateTime: '2024-01-01T09:00:00', TimeZone: 'Europe/London' });
    expect(s.ScheduledEndDateTime).toEqual({ DateTime: '2024-01-05T17:00:00', TimeZone: 'Europe/London' });
  });

  it('scheduled without start/end omits the datetimes (default UTC path)', async () => {
    mPatch.mockResolvedValue({});
    await setAutomaticReplies({ status: 'Scheduled', internalMessage: 'OOO' });
    const s = (mPatch.mock.calls[0][1] as any).AutomaticRepliesSetting;
    expect(s.ScheduledStartDateTime).toBeUndefined();
    expect(s.ScheduledEndDateTime).toBeUndefined();
  });
});
