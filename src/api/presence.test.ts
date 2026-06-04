import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('./client.js', () => ({ owaPost: vi.fn() }));
vi.mock('./calendar.js', () => ({ getSchedule: vi.fn() }));

import { owaPost } from './client.js';
import { getSchedule } from './calendar.js';
import { getMailTips, getAvailability } from './presence.js';

const mPost = vi.mocked(owaPost);
const mSchedule = vi.mocked(getSchedule);

beforeEach(() => vi.clearAllMocks());

describe('getMailTips', () => {
  it('returns value array', async () => {
    mPost.mockResolvedValue({ value: [{ EmailAddress: { Address: 'a@b.com' } }] });
    const res = await getMailTips(['a@b.com']);
    expect(res).toEqual([{ EmailAddress: { Address: 'a@b.com' } }]);
    const body = mPost.mock.calls[0][1] as any;
    expect(mPost.mock.calls[0][0]).toBe('/getmailtips');
    expect(body.EmailAddresses).toEqual(['a@b.com']);
    expect(body.MailTipsOptions).toBe('automaticReplies');
  });
  it('returns [] when value missing', async () => {
    mPost.mockResolvedValue({});
    const res = await getMailTips(['a@b.com']);
    expect(res).toEqual([]);
  });
});

describe('getAvailability', () => {
  it('combines mailtips and schedule into availability', async () => {
    mPost.mockResolvedValue({ value: [
      { EmailAddress: { Address: 'A@B.com', Name: 'Alice' }, AutomaticReplies: { Message: '  Out today  ', ScheduledStartTime: { DateTime: 'S', TimeZone: 'UTC' }, ScheduledEndTime: { DateTime: 'E', TimeZone: 'UTC' } } },
      { EmailAddress: {}, AutomaticReplies: { Message: 'x' } },
    ] });
    mSchedule.mockResolvedValue([
      { ScheduleId: 'a@b.com', AvailabilityView: '02', WorkingHours: { DaysOfWeek: ['Monday'], StartTime: '09:00', EndTime: '17:00', TimeZone: { Name: 'GMT' } } },
    ] as any);

    const res = await getAvailability(['a@b.com'], 4);
    expect(res).toHaveLength(1);
    const info = res[0];
    expect(info.email).toBe('a@b.com');
    expect(info.displayName).toBe('Alice');
    expect(info.status).toBe('Free');
    expect(typeof info.statusChangesAt).toBe('string');
    expect(info.workingHours).toEqual({ daysOfWeek: ['Monday'], startTime: '09:00', endTime: '17:00', timeZone: 'GMT' });
    expect(info.outOfOffice).toEqual({ isActive: true, message: 'Out today', scheduledStart: 'S', scheduledEnd: 'E' });
  });

  it('handles inactive OOO, unknown status, no working hours and no status change', async () => {
    mPost.mockResolvedValue({ value: [
      { EmailAddress: { Address: 'a@b.com' }, AutomaticReplies: { Message: '   ' } },
    ] });
    mSchedule.mockResolvedValue([
      { ScheduleId: 'a@b.com', AvailabilityView: '99' },
    ] as any);

    const res = await getAvailability(['a@b.com']);
    const info = res[0];
    expect(info.displayName).toBeUndefined();
    expect(info.status).toBe('Unknown');
    expect(info.statusChangesAt).toBeUndefined();
    expect(info.workingHours).toBeUndefined();
    expect(info.outOfOffice).toEqual({ isActive: false });
  });

  it('handles empty availability view and missing schedule entry', async () => {
    mPost.mockResolvedValue({ value: [] });
    mSchedule.mockResolvedValue([
      { ScheduleId: 'other@b.com', AvailabilityView: '' },
    ] as any);

    const res = await getAvailability(['a@b.com']);
    const info = res[0];
    expect(info.status).toBeUndefined();
    expect(info.outOfOffice).toBeUndefined();
  });

  it('records per-section errors when both calls reject', async () => {
    mPost.mockRejectedValue(new Error('tips failed'));
    mSchedule.mockRejectedValue('schedule failed');

    const res = await getAvailability(['a@b.com']);
    const info = res[0];
    expect(info.scheduleError).toBe('schedule failed');
    expect(info.outOfOfficeError).toBe('tips failed');
    expect(info.status).toBeUndefined();
    expect(info.outOfOffice).toBeUndefined();
  });

  it('does not set status change when view never changes', async () => {
    mPost.mockResolvedValue({ value: [] });
    mSchedule.mockResolvedValue([
      { ScheduleId: 'a@b.com', AvailabilityView: '222' },
    ] as any);
    const res = await getAvailability(['a@b.com']);
    expect(res[0].status).toBe('Busy');
    expect(res[0].statusChangesAt).toBeUndefined();
  });
});

describe('getAvailability schedule without AvailabilityView', () => {
  it('treats missing view as empty', async () => {
    vi.mocked(owaPost).mockResolvedValue({ value: [] });
    vi.mocked(getSchedule).mockResolvedValue([{ ScheduleId: 'a@b.com' }] as any);
    const res = await getAvailability(['a@b.com']);
    expect(res[0].status).toBeUndefined();
  });
});
