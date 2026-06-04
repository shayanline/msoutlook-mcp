import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('./client.js', () => ({
  owaGet: vi.fn(),
  owaPost: vi.fn(),
  owaPatch: vi.fn(),
  owaDelete: vi.fn(),
}));

import { owaGet, owaPost, owaPatch, owaDelete } from './client.js';
import {
  listEvents, getEvent, createEvent, updateEvent, deleteEvent, respondToEvent,
  searchEvents, listCalendars, getSchedule, findMeetingTimes, cancelEvent, forwardEvent,
} from './calendar.js';

const mGet = vi.mocked(owaGet);
const mPost = vi.mocked(owaPost);
const mPatch = vi.mocked(owaPatch);
const mDelete = vi.mocked(owaDelete);

beforeEach(() => vi.clearAllMocks());

describe('listEvents', () => {
  it('uses calendarview and defaults', async () => {
    mGet.mockResolvedValue({ value: [{ Id: 'e1' }] });
    const res = await listEvents();
    expect(res).toEqual([{ Id: 'e1' }]);
    const [path, params] = mGet.mock.calls[0] as [string, Record<string, string>];
    expect(path).toBe('/calendarview');
    expect(params['$top']).toBe('50');
    expect(params['$orderby']).toBe('start/dateTime asc');
    expect(typeof params.startDateTime).toBe('string');
    expect(typeof params.endDateTime).toBe('string');
    expect(params['$filter']).toBeUndefined();
  });

  it('uses calendar-specific path, custom dates, select and filter', async () => {
    mGet.mockResolvedValue({ value: [] });
    await listEvents({ calendarId: 'cal1', startDateTime: 'S', endDateTime: 'E', top: 5, select: ['Id'], filter: 'x eq 1' });
    const [path, params] = mGet.mock.calls[0] as [string, Record<string, string>];
    expect(path).toBe('/calendars/cal1/calendarview');
    expect(params.startDateTime).toBe('S');
    expect(params.endDateTime).toBe('E');
    expect(params['$top']).toBe('5');
    expect(params['$select']).toBe('Id');
    expect(params['$filter']).toBe('x eq 1');
  });
});

describe('getEvent', () => {
  it('fetches by id', async () => {
    mGet.mockResolvedValue({ Id: 'e1' });
    await getEvent('e1');
    expect(mGet).toHaveBeenCalledWith('/events/e1');
  });
});

describe('createEvent', () => {
  it('creates with defaults', async () => {
    mPost.mockResolvedValue({ Id: 'e1' });
    await createEvent({ subject: 'Sync', start: '2024-01-01T09:00', end: '2024-01-01T10:00' });
    const body = mPost.mock.calls[0][1] as any;
    expect(mPost.mock.calls[0][0]).toBe('/events');
    expect(body.Body).toEqual({ ContentType: 'Text', Content: '' });
    expect(body.Start).toEqual({ DateTime: '2024-01-01T09:00', TimeZone: 'UTC' });
    expect(body.Location).toBeUndefined();
    expect(body.Attendees).toBeUndefined();
    expect(body.IsOnlineMeeting).toBe(false);
    expect(body.Importance).toBe('Normal');
    expect(body.IsAllDay).toBe(false);
  });

  it('creates with location, attendees, online, importance, allDay and tz', async () => {
    mPost.mockResolvedValue({ Id: 'e2' });
    await createEvent({
      subject: 'S', body: 'desc', bodyType: 'HTML', start: 's', end: 'e', timeZone: 'Europe/London',
      location: 'Room 1', isOnlineMeeting: true, importance: 'High', isAllDay: true,
      attendees: [{ email: 'a@b.com', name: 'A', type: 'Optional' }, { email: 'c@d.com' }],
    });
    const body = mPost.mock.calls[0][1] as any;
    expect(body.Body).toEqual({ ContentType: 'HTML', Content: 'desc' });
    expect(body.Start.TimeZone).toBe('Europe/London');
    expect(body.Location).toEqual({ DisplayName: 'Room 1' });
    expect(body.Attendees).toEqual([
      { EmailAddress: { Address: 'a@b.com', Name: 'A' }, Type: 'Optional' },
      { EmailAddress: { Address: 'c@d.com', Name: 'c@d.com' }, Type: 'Required' },
    ]);
    expect(body.IsOnlineMeeting).toBe(true);
    expect(body.IsAllDay).toBe(true);
  });
});

describe('updateEvent', () => {
  it('builds patch from all provided fields', async () => {
    mPatch.mockResolvedValue({ Id: 'e1' });
    await updateEvent('e1', {
      subject: 'New', body: 'b', bodyType: 'HTML', start: 's', end: 'e', timeZone: 'TZ',
      location: 'L', isOnlineMeeting: true,
    });
    const body = mPatch.mock.calls[0][1] as any;
    expect(mPatch.mock.calls[0][0]).toBe('/events/e1');
    expect(body.Subject).toBe('New');
    expect(body.Body).toEqual({ ContentType: 'HTML', Content: 'b' });
    expect(body.Start).toEqual({ DateTime: 's', TimeZone: 'TZ' });
    expect(body.End).toEqual({ DateTime: 'e', TimeZone: 'TZ' });
    expect(body.Location).toEqual({ DisplayName: 'L' });
    expect(body.IsOnlineMeeting).toBe(true);
  });
  it('defaults body content type to Text and tz to UTC', async () => {
    mPatch.mockResolvedValue({ Id: 'e1' });
    await updateEvent('e1', { body: 'x', start: 's' });
    const body = mPatch.mock.calls[0][1] as any;
    expect(body.Body).toEqual({ ContentType: 'Text', Content: 'x' });
    expect(body.Start.TimeZone).toBe('UTC');
  });
  it('sends empty patch when nothing set', async () => {
    mPatch.mockResolvedValue({ Id: 'e1' });
    await updateEvent('e1', {});
    expect(mPatch).toHaveBeenCalledWith('/events/e1', {});
  });
});

describe('deleteEvent', () => {
  it('deletes', async () => {
    mDelete.mockResolvedValue(undefined);
    await deleteEvent('e1');
    expect(mDelete).toHaveBeenCalledWith('/events/e1');
  });
});

describe('respondToEvent', () => {
  it('responds with comment', async () => {
    mPost.mockResolvedValue(undefined);
    await respondToEvent('e1', 'accept', 'sure');
    expect(mPost).toHaveBeenCalledWith('/events/e1/accept', { Comment: 'sure', SendResponse: true });
  });
  it('responds with default empty comment', async () => {
    mPost.mockResolvedValue(undefined);
    await respondToEvent('e1', 'decline');
    expect(mPost).toHaveBeenCalledWith('/events/e1/decline', { Comment: '', SendResponse: true });
  });
});

describe('searchEvents', () => {
  it('delegates to listEvents with contains filter and escaped quotes', async () => {
    mGet.mockResolvedValue({ value: [] });
    await searchEvents("O'Brien", 'F', 'T');
    const [path, params] = mGet.mock.calls[0] as [string, Record<string, string>];
    expect(path).toBe('/calendarview');
    expect(params['$filter']).toBe("contains(subject,'O''Brien')");
    expect(params.startDateTime).toBe('F');
    expect(params.endDateTime).toBe('T');
    expect(params['$top']).toBe('25');
  });
});

describe('listCalendars', () => {
  it('lists calendars', async () => {
    mGet.mockResolvedValue({ value: [{ Id: 'c1' }] });
    const res = await listCalendars();
    expect(res).toEqual([{ Id: 'c1' }]);
    expect(mGet.mock.calls[0][0]).toBe('/calendars');
  });
});

describe('getSchedule', () => {
  it('posts schedule request with default interval', async () => {
    mPost.mockResolvedValue({ value: [{ ScheduleId: 'a@b.com' }] });
    const start = new Date('2024-01-01T00:00:00Z');
    const end = new Date('2024-01-01T08:00:00Z');
    const res = await getSchedule(['a@b.com'], start, end);
    expect(res).toEqual([{ ScheduleId: 'a@b.com' }]);
    const body = mPost.mock.calls[0][1] as any;
    expect(mPost.mock.calls[0][0]).toBe('/calendar/getschedule');
    expect(body.Schedules).toEqual(['a@b.com']);
    expect(body.StartTime).toEqual({ DateTime: start.toISOString(), TimeZone: 'UTC' });
    expect(body.AvailabilityViewInterval).toBe(30);
  });
  it('returns [] when value missing and uses custom interval', async () => {
    mPost.mockResolvedValue({});
    const res = await getSchedule(['a@b.com'], new Date(), new Date(), 15);
    expect(res).toEqual([]);
    expect((mPost.mock.calls[0][1] as any).AvailabilityViewInterval).toBe(15);
  });
});

describe('findMeetingTimes', () => {
  it('uses defaults and maps suggestions', async () => {
    mPost.mockResolvedValue({ MeetingTimeSuggestions: [{ Confidence: 100 }] });
    const res = await findMeetingTimes({ attendees: ['a@b.com'] });
    expect(res.suggestions).toEqual([{ Confidence: 100 }]);
    expect(res.emptyReason).toBeUndefined();
    const body = mPost.mock.calls[0][1] as any;
    expect(mPost.mock.calls[0][0]).toBe('/findmeetingtimes');
    expect(body.Attendees).toEqual([{ Type: 'Required', EmailAddress: { Address: 'a@b.com' } }]);
    expect(body.MeetingDuration).toBe('PT30M');
    expect(body.MaxCandidates).toBe(5);
  });
  it('honours custom options and returns empty reason', async () => {
    mPost.mockResolvedValue({ EmptySuggestionsReason: 'AttendeesUnavailable' });
    const res = await findMeetingTimes({
      attendees: ['a@b.com'], durationMinutes: 60, start: 'S', end: 'E', maxCandidates: 3, timeZone: 'TZ',
    });
    expect(res.suggestions).toEqual([]);
    expect(res.emptyReason).toBe('AttendeesUnavailable');
    const body = mPost.mock.calls[0][1] as any;
    expect(body.MeetingDuration).toBe('PT60M');
    expect(body.MaxCandidates).toBe(3);
    expect(body.TimeConstraint.Timeslots[0].Start).toEqual({ DateTime: 'S', TimeZone: 'TZ' });
    expect(body.TimeConstraint.Timeslots[0].End).toEqual({ DateTime: 'E', TimeZone: 'TZ' });
  });
  it('treats empty-string EmptySuggestionsReason as undefined', async () => {
    mPost.mockResolvedValue({ MeetingTimeSuggestions: [], EmptySuggestionsReason: '' });
    const res = await findMeetingTimes({ attendees: ['a@b.com'] });
    expect(res.emptyReason).toBeUndefined();
  });
});

describe('cancelEvent', () => {
  it('cancels with comment', async () => {
    mPost.mockResolvedValue(undefined);
    await cancelEvent('e1', 'sorry');
    expect(mPost).toHaveBeenCalledWith('/events/e1/cancel', { Comment: 'sorry' });
  });
  it('cancels with default comment', async () => {
    mPost.mockResolvedValue(undefined);
    await cancelEvent('e1');
    expect(mPost).toHaveBeenCalledWith('/events/e1/cancel', { Comment: '' });
  });
});

describe('forwardEvent', () => {
  it('forwards with comment', async () => {
    mPost.mockResolvedValue(undefined);
    await forwardEvent('e1', ['a@b.com'], 'fyi');
    const body = mPost.mock.calls[0][1] as any;
    expect(body.Comment).toBe('fyi');
    expect(body.ToRecipients).toEqual([{ EmailAddress: { Address: 'a@b.com' } }]);
  });
  it('forwards with default comment', async () => {
    mPost.mockResolvedValue(undefined);
    await forwardEvent('e1', ['a@b.com']);
    expect((mPost.mock.calls[0][1] as any).Comment).toBe('');
  });
});
