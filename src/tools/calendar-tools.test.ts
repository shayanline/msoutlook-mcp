import { describe, it, expect, vi, beforeEach } from 'vitest';
import * as cal from '../api/calendar.js';
import { registerCalendarTools } from './calendar-tools.js';

vi.mock('../api/calendar.js', () => ({
  listEvents: vi.fn(), getEvent: vi.fn(), createEvent: vi.fn(), updateEvent: vi.fn(),
  deleteEvent: vi.fn(), respondToEvent: vi.fn(), searchEvents: vi.fn(), listCalendars: vi.fn(),
  getSchedule: vi.fn(), findMeetingTimes: vi.fn(), cancelEvent: vi.fn(), forwardEvent: vi.fn(),
}));

type Reg = { schema: any; handler: (a: any) => Promise<any> };
let tools: Map<string, Reg>;
function setup() {
  tools = new Map();
  const server = { tool: (n: string, _d: string, s: any, h: any) => tools.set(n, { schema: s, handler: h }) } as any;
  registerCalendarTools(server);
}
const text = (r: any) => r.content[0].text;

const fullEvent = {
  Id: 'e1', Subject: 'Sync', Start: { DateTime: '2024-01-01T10:00', TimeZone: 'UTC' }, End: { DateTime: '2024-01-01T11:00' },
  Location: { DisplayName: 'Room A' }, Organizer: { EmailAddress: { Address: 'org@x.com' } },
  Attendees: [
    { EmailAddress: { Name: 'Jane', Address: 'jane@x.com' }, Type: 'Required', Status: { Response: 'accepted' } },
    { EmailAddress: { Address: 'bob@x.com' } },
  ],
  IsOnlineMeeting: true, OnlineMeetingUrl: 'http://meet', ResponseStatus: { Response: 'organizer' },
  WebLink: 'http://w', Body: { ContentType: 'HTML', Content: 'body' },
};
const leanEvent = {
  Id: 'e2', Subject: 'X', IsOnlineMeeting: true, BodyPreview: 'p'.repeat(300),
};

beforeEach(() => { vi.clearAllMocks(); setup(); });

describe('list_events', () => {
  it('lists lean event with defaults', async () => {
    vi.mocked(cal.listEvents).mockResolvedValue([leanEvent] as any);
    const r = await tools.get('outlook_list_events')!.handler({});
    expect(cal.listEvents).toHaveBeenCalledWith({ startDateTime: undefined, endDateTime: undefined, top: 50, calendarId: undefined });
    expect(text(r)).toContain('Unknown');
    expect(text(r)).toContain('Online Meeting: Yes');
    expect(text(r)).toContain('Preview:');
  });
  it('empty', async () => {
    vi.mocked(cal.listEvents).mockResolvedValue([] as any);
    expect(text(await tools.get('outlook_list_events')!.handler({ start_date: 's', end_date: 'e', top: 5, calendar_id: 'c' }))).toBe('No events found.');
  });
});

describe('get_event', () => {
  it('full details', async () => {
    vi.mocked(cal.getEvent).mockResolvedValue(fullEvent as any);
    const r = text(await tools.get('outlook_get_event')!.handler({ id: 'e1' }));
    expect(r).toContain('Location: Room A');
    expect(r).toContain('Organizer: org@x.com');
    expect(r).toContain('Online Meeting: http://meet');
    expect(r).toContain('Attendees:');
    expect(r).toContain('Jane (Required) — accepted');
    expect(r).toContain('bob@x.com (Required) — None');
    expect(r).toContain('Body (HTML):');
  });
});

describe('create/update/delete/respond', () => {
  it('create_event', async () => {
    vi.mocked(cal.createEvent).mockResolvedValue({ Id: 'e', Subject: 's', Start: { DateTime: 'd' } } as any);
    const r = await tools.get('outlook_create_event')!.handler({
      subject: 's', start: 'a', end: 'b', time_zone: 'UTC', location: 'L', body: 'B',
      attendees: [{ email: 'a@x.com' }], is_online_meeting: true, is_all_day: false, importance: 'High',
    });
    expect(text(r)).toContain('Event created.');
    expect(cal.createEvent).toHaveBeenCalled();
  });
  it('create_event maps recurrence and reminder args to camelCase options', async () => {
    vi.mocked(cal.createEvent).mockResolvedValue({ Id: 'e', Subject: 's', Start: { DateTime: 'd' } } as any);
    await tools.get('outlook_create_event')!.handler({
      subject: 's', start: '2026-06-10T09:00:00', end: '2026-06-10T09:30:00', time_zone: 'Europe/London',
      recurrence: {
        pattern: 'relativeMonthly', days_of_week: ['friday'], index: 'last',
        range: { type: 'numbered', count: 6, start_date: '2026-06-26', time_zone: 'Europe/London' },
      },
      reminder_minutes_before_start: 15, is_reminder_on: true,
      show_as: 'busy', categories: ['Planning'], is_private: true,
    });
    const opts = vi.mocked(cal.createEvent).mock.calls[0][0] as any;
    expect(opts.recurrence).toEqual({
      pattern: 'relativeMonthly', interval: undefined, daysOfWeek: ['friday'], dayOfMonth: undefined,
      month: undefined, index: 'last', firstDayOfWeek: undefined,
      range: { type: 'numbered', startDate: '2026-06-26', endDate: undefined, numberOfOccurrences: 6, timeZone: 'Europe/London' },
    });
    expect(opts.reminderMinutesBeforeStart).toBe(15);
    expect(opts.isReminderOn).toBe(true);
    expect(opts.showAs).toBe('busy');
    expect(opts.categories).toEqual(['Planning']);
    expect(opts.isPrivate).toBe(true);
  });
  it('update_event', async () => {
    vi.mocked(cal.updateEvent).mockResolvedValue({ Id: 'e', Subject: 's' } as any);
    const r = await tools.get('outlook_update_event')!.handler({ id: 'e', subject: 's', start: 'a', end: 'b', time_zone: 'UTC', location: 'L', body: 'B' });
    expect(text(r)).toContain('Event updated.');
    expect(cal.updateEvent).toHaveBeenCalledWith('e', {
      subject: 's', start: 'a', end: 'b', timeZone: 'UTC', location: 'L', body: 'B',
      recurrence: undefined, reminderMinutesBeforeStart: undefined, isReminderOn: undefined,
      showAs: undefined, categories: undefined, isPrivate: undefined,
    });
  });
  it('delete_event', async () => {
    vi.mocked(cal.deleteEvent).mockResolvedValue(undefined as any);
    expect(text(await tools.get('outlook_delete_event')!.handler({ id: 'e' }))).toBe('Event deleted.');
  });
  it('respond_to_event', async () => {
    vi.mocked(cal.respondToEvent).mockResolvedValue(undefined as any);
    expect(text(await tools.get('outlook_respond_to_event')!.handler({ id: 'e', response: 'accept', comment: 'ok' }))).toContain('accept');
    expect(cal.respondToEvent).toHaveBeenCalledWith('e', 'accept', 'ok');
  });
});

describe('search_events', () => {
  it('found', async () => {
    vi.mocked(cal.searchEvents).mockResolvedValue([fullEvent] as any);
    await tools.get('outlook_search_events')!.handler({ query: 'q', from_date: 'f', to_date: 't' });
    expect(cal.searchEvents).toHaveBeenCalledWith('q', 'f', 't');
  });
  it('empty', async () => {
    vi.mocked(cal.searchEvents).mockResolvedValue([] as any);
    expect(text(await tools.get('outlook_search_events')!.handler({ query: 'q' }))).toBe('No events found.');
  });
});

describe('list_calendars', () => {
  it('lists with flags', async () => {
    vi.mocked(cal.listCalendars).mockResolvedValue([
      { Name: 'Cal', Id: 'c1', IsDefaultCalendar: true, CanEdit: true },
      { Name: 'Other', Id: 'c2' },
    ] as any);
    const r = text(await tools.get('outlook_list_calendars')!.handler({}));
    expect(r).toContain('Cal (ID: c1) [Default] [Editable]');
    expect(r).toContain('Other (ID: c2)');
  });
  it('empty', async () => {
    vi.mocked(cal.listCalendars).mockResolvedValue([] as any);
    expect(text(await tools.get('outlook_list_calendars')!.handler({}))).toBe('No calendars found.');
  });
});

describe('get_schedule', () => {
  it('formats busy + working hours and free', async () => {
    vi.mocked(cal.getSchedule).mockResolvedValue([
      {
        ScheduleId: 'a@x.com',
        ScheduleItems: [
          { Status: 'Busy', Start: { DateTime: 's1' }, End: { DateTime: 'e1' }, Subject: 'Mtg' },
          { Status: 'Tentative' },
          { Status: 'Free' },
        ],
        WorkingHours: { StartTime: '09:00', EndTime: '17:00', TimeZone: { Name: 'UTC' } },
      },
      { ScheduleId: 'b@x.com', ScheduleItems: [], WorkingHours: {} },
    ] as any);
    const r = text(await tools.get('outlook_get_schedule')!.handler({ emails: ['a@x.com'], start: '2024-01-01', end: '2024-01-02', interval_minutes: 15 }));
    expect(r).toContain('Busy: s1 to e1 (Mtg)');
    expect(r).toContain('Tentative: ? to ?');
    expect(r).toContain('Working hours: 09:00 to 17:00 UTC');
    expect(r).toContain('No busy blocks');
  });
  it('default window + empty schedules', async () => {
    vi.mocked(cal.getSchedule).mockResolvedValue([] as any);
    const r = await tools.get('outlook_get_schedule')!.handler({ emails: ['a@x.com'] });
    expect(text(r)).toBe('No schedule information returned.');
    const args = vi.mocked(cal.getSchedule).mock.calls[0];
    expect(args[0]).toEqual(['a@x.com']);
    expect(args[3]).toBe(30);
  });
});

describe('find_meeting_times', () => {
  it('suggestions present', async () => {
    vi.mocked(cal.findMeetingTimes).mockResolvedValue({
      suggestions: [
        { Confidence: 90, MeetingTimeSlot: { Start: { DateTime: 's', TimeZone: 'UTC' }, End: { DateTime: 'e' } }, AttendeeAvailability: [{ Attendee: { EmailAddress: { Address: 'a@x.com' } }, Availability: 'Free' }] },
        { MeetingTimeSlot: undefined },
      ],
      emptyReason: undefined,
    } as any);
    const r = text(await tools.get('outlook_find_meeting_times')!.handler({ attendees: ['a@x.com'], duration_minutes: 60, start: 's', end: 'e', max_candidates: 3, time_zone: 'UTC' }));
    expect(r).toContain('confidence 90%');
    expect(r).toContain('a@x.com: Free');
    expect(r).toContain('confidence ?%');
  });
  it('no suggestions with reason', async () => {
    vi.mocked(cal.findMeetingTimes).mockResolvedValue({ suggestions: [], emptyReason: 'busy' } as any);
    expect(text(await tools.get('outlook_find_meeting_times')!.handler({ attendees: ['a@x.com'] }))).toContain('Reason: busy');
  });
  it('no suggestions no reason', async () => {
    vi.mocked(cal.findMeetingTimes).mockResolvedValue({ suggestions: [], emptyReason: undefined } as any);
    expect(text(await tools.get('outlook_find_meeting_times')!.handler({ attendees: ['a@x.com'] }))).toBe('No meeting times found.');
  });
});

describe('cancel/forward', () => {
  it('cancel_event', async () => {
    vi.mocked(cal.cancelEvent).mockResolvedValue(undefined as any);
    expect(text(await tools.get('outlook_cancel_event')!.handler({ id: 'e', comment: 'c' }))).toContain('cancelled');
  });
  it('forward_event', async () => {
    vi.mocked(cal.forwardEvent).mockResolvedValue(undefined as any);
    expect(text(await tools.get('outlook_forward_event')!.handler({ id: 'e', to: ['a@x.com'], comment: 'c' }))).toContain('forwarded to a@x.com');
  });
});
