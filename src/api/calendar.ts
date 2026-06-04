/**
 * Calendar API — events, meetings, and availability.
 */

import { owaGet, owaPost, owaPatch, owaDelete } from './client.js';
import type { ODataResponse } from './mail.js';
import type { Recipient } from './mail.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

export interface DateTimeTimeZone {
  DateTime: string;
  TimeZone: string;
}

export interface Attendee {
  EmailAddress: { Name?: string; Address: string };
  Type?: 'Required' | 'Optional' | 'Resource';
  Status?: { Response: string; Time: string };
}

export interface CalendarEvent {
  Id: string;
  Subject: string;
  BodyPreview?: string;
  Body?: { ContentType: string; Content: string };
  Start: DateTimeTimeZone;
  End: DateTimeTimeZone;
  Location?: { DisplayName: string };
  Organizer?: Recipient;
  Attendees?: Attendee[];
  IsAllDay?: boolean;
  IsCancelled?: boolean;
  IsOnlineMeeting?: boolean;
  OnlineMeetingUrl?: string;
  Recurrence?: unknown;
  Importance?: string;
  Sensitivity?: string;
  ShowAs?: string;
  WebLink?: string;
  ResponseStatus?: { Response: string; Time: string };
}

export interface Calendar {
  Id: string;
  Name: string;
  Color?: string;
  IsDefaultCalendar?: boolean;
  CanEdit?: boolean;
}

// ─────────────────────────────────────────────────────────────────────────────
// List events
// ─────────────────────────────────────────────────────────────────────────────

export interface ListEventsOptions {
  calendarId?: string;
  startDateTime?: string;
  endDateTime?: string;
  top?: number;
  filter?: string;
  select?: string[];
}

export async function listEvents(opts: ListEventsOptions = {}): Promise<CalendarEvent[]> {
  const path = opts.calendarId ? `/calendars/${opts.calendarId}/calendarview` : '/calendarview';

  const now = new Date();
  const weekOut = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);

  const params: Record<string, string> = {
    startDateTime: opts.startDateTime ?? now.toISOString(),
    endDateTime: opts.endDateTime ?? weekOut.toISOString(),
    '$top': String(opts.top ?? 50),
    '$orderby': 'start/dateTime asc',
    '$select': (opts.select ?? [
      'Id', 'Subject', 'BodyPreview', 'Start', 'End', 'Location',
      'Organizer', 'Attendees', 'IsAllDay', 'IsCancelled', 'IsOnlineMeeting',
      'OnlineMeetingUrl', 'ResponseStatus', 'WebLink',
    ]).join(','),
  };

  if (opts.filter) params['$filter'] = opts.filter;

  const res = await owaGet<ODataResponse<CalendarEvent>>(path, params);
  return res.value;
}

// ─────────────────────────────────────────────────────────────────────────────
// Get event
// ─────────────────────────────────────────────────────────────────────────────

export async function getEvent(id: string): Promise<CalendarEvent> {
  return owaGet<CalendarEvent>(`/events/${id}`);
}

// ─────────────────────────────────────────────────────────────────────────────
// Create event
// ─────────────────────────────────────────────────────────────────────────────

export interface CreateEventOptions {
  subject: string;
  body?: string;
  bodyType?: 'Text' | 'HTML';
  start: string;
  end: string;
  timeZone?: string;
  location?: string;
  attendees?: Array<{ email: string; name?: string; type?: 'Required' | 'Optional' }>;
  isOnlineMeeting?: boolean;
  importance?: 'Low' | 'Normal' | 'High';
  isAllDay?: boolean;
}

export async function createEvent(opts: CreateEventOptions): Promise<CalendarEvent> {
  const tz = opts.timeZone ?? 'UTC';
  const attendees = opts.attendees?.map(a => ({
    EmailAddress: { Address: a.email, Name: a.name ?? a.email },
    Type: a.type ?? 'Required',
  }));

  return owaPost<CalendarEvent>('/events', {
    Subject: opts.subject,
    Body: { ContentType: opts.bodyType ?? 'Text', Content: opts.body ?? '' },
    Start: { DateTime: opts.start, TimeZone: tz },
    End: { DateTime: opts.end, TimeZone: tz },
    ...(opts.location ? { Location: { DisplayName: opts.location } } : {}),
    ...(attendees?.length ? { Attendees: attendees } : {}),
    IsOnlineMeeting: opts.isOnlineMeeting ?? false,
    Importance: opts.importance ?? 'Normal',
    IsAllDay: opts.isAllDay ?? false,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Update event
// ─────────────────────────────────────────────────────────────────────────────

export async function updateEvent(id: string, updates: Partial<CreateEventOptions>): Promise<CalendarEvent> {
  const tz = updates.timeZone ?? 'UTC';
  const body: Record<string, unknown> = {};

  if (updates.subject !== undefined) body.Subject = updates.subject;
  if (updates.body !== undefined) {
    body.Body = { ContentType: updates.bodyType ?? 'Text', Content: updates.body };
  }
  if (updates.start !== undefined) body.Start = { DateTime: updates.start, TimeZone: tz };
  if (updates.end !== undefined) body.End = { DateTime: updates.end, TimeZone: tz };
  if (updates.location !== undefined) body.Location = { DisplayName: updates.location };
  if (updates.isOnlineMeeting !== undefined) body.IsOnlineMeeting = updates.isOnlineMeeting;

  return owaPatch<CalendarEvent>(`/events/${id}`, body);
}

// ─────────────────────────────────────────────────────────────────────────────
// Delete event
// ─────────────────────────────────────────────────────────────────────────────

export async function deleteEvent(id: string): Promise<void> {
  await owaDelete(`/events/${id}`);
}

// ─────────────────────────────────────────────────────────────────────────────
// Respond to event (accept/decline/tentative)
// ─────────────────────────────────────────────────────────────────────────────

export type EventResponse = 'accept' | 'decline' | 'tentativelyAccept';

export async function respondToEvent(id: string, response: EventResponse, comment?: string): Promise<void> {
  await owaPost(`/events/${id}/${response}`, {
    Comment: comment ?? '',
    SendResponse: true,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Search events
// ─────────────────────────────────────────────────────────────────────────────

export async function searchEvents(query: string, fromDate?: string, toDate?: string): Promise<CalendarEvent[]> {
  return listEvents({
    startDateTime: fromDate,
    endDateTime: toDate,
    filter: `contains(subject,'${query.replace(/'/g, "''")}')`,
    top: 25,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Calendars
// ─────────────────────────────────────────────────────────────────────────────

export async function listCalendars(): Promise<Calendar[]> {
  const res = await owaGet<ODataResponse<Calendar>>('/calendars', {
    '$select': 'Id,Name,Color,IsDefaultCalendar,CanEdit',
  });
  return res.value;
}

// ─────────────────────────────────────────────────────────────────────────────
// Free/busy schedule
// ─────────────────────────────────────────────────────────────────────────────

export interface ScheduleItem {
  Status?: string;
  Start?: DateTimeTimeZone;
  End?: DateTimeTimeZone;
  Subject?: string;
  Location?: string;
}

export interface ScheduleInformation {
  ScheduleId: string;
  AvailabilityView?: string;
  ScheduleItems?: ScheduleItem[];
  WorkingHours?: {
    DaysOfWeek?: string[];
    StartTime?: string;
    EndTime?: string;
    TimeZone?: { Name?: string };
  };
}

/** Free/busy for a batch of addresses over a window (one digit per interval). */
export async function getSchedule(
  emails: string[],
  start: Date,
  end: Date,
  intervalMinutes = 30,
): Promise<ScheduleInformation[]> {
  const res = await owaPost<ODataResponse<ScheduleInformation>>('/calendar/getschedule', {
    Schedules: emails,
    StartTime: { DateTime: start.toISOString(), TimeZone: 'UTC' },
    EndTime: { DateTime: end.toISOString(), TimeZone: 'UTC' },
    AvailabilityViewInterval: intervalMinutes,
  });
  return res.value ?? [];
}

// ─────────────────────────────────────────────────────────────────────────────
// Find meeting times
// ─────────────────────────────────────────────────────────────────────────────

export interface MeetingTimeSuggestion {
  Confidence?: number;
  OrganizerAvailability?: string;
  MeetingTimeSlot?: { Start: DateTimeTimeZone; End: DateTimeTimeZone };
  AttendeeAvailability?: Array<{ Availability: string; Attendee: { EmailAddress: { Address: string } } }>;
}

export interface FindMeetingTimesOptions {
  attendees: string[];
  durationMinutes?: number;
  start?: string;
  end?: string;
  maxCandidates?: number;
  timeZone?: string;
}

export async function findMeetingTimes(opts: FindMeetingTimesOptions): Promise<{
  suggestions: MeetingTimeSuggestion[];
  emptyReason?: string;
}> {
  const tz = opts.timeZone ?? 'UTC';
  const now = new Date();
  const start = opts.start ?? now.toISOString();
  const end = opts.end ?? new Date(now.getTime() + 5 * 24 * 60 * 60 * 1000).toISOString();
  const minutes = opts.durationMinutes ?? 30;

  const res = await owaPost<{ MeetingTimeSuggestions?: MeetingTimeSuggestion[]; EmptySuggestionsReason?: string }>(
    '/findmeetingtimes',
    {
      Attendees: opts.attendees.map(addr => ({ Type: 'Required', EmailAddress: { Address: addr } })),
      TimeConstraint: {
        Timeslots: [{ Start: { DateTime: start, TimeZone: tz }, End: { DateTime: end, TimeZone: tz } }],
      },
      MeetingDuration: `PT${minutes}M`,
      MaxCandidates: opts.maxCandidates ?? 5,
    },
  );
  return { suggestions: res.MeetingTimeSuggestions ?? [], emptyReason: res.EmptySuggestionsReason || undefined };
}

// ─────────────────────────────────────────────────────────────────────────────
// Cancel / forward event
// ─────────────────────────────────────────────────────────────────────────────

/** Cancel an event you organize, notifying attendees. */
export async function cancelEvent(id: string, comment?: string): Promise<void> {
  await owaPost(`/events/${id}/cancel`, { Comment: comment ?? '' });
}

/** Forward a meeting invite to additional people. */
export async function forwardEvent(id: string, to: string[], comment?: string): Promise<void> {
  await owaPost(`/events/${id}/forward`, {
    ToRecipients: to.map(addr => ({ EmailAddress: { Address: addr } })),
    Comment: comment ?? '',
  });
}
