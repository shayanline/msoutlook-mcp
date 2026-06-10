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
  ReminderMinutesBeforeStart?: number;
  IsReminderOn?: boolean;
  Categories?: string[];
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
// Recurrence
//
// Maps a clean, friendly input shape to the OWA REST v2 `Recurrence` property
// (a PatternedRecurrence: Pattern + Range), the same model Microsoft Graph
// exposes as `patternedRecurrence`. Input values use Graph-style camelCase
// (e.g. 'absoluteMonthly', 'monday', 'first') and are translated to the
// PascalCase enum values OWA expects ('AbsoluteMonthly', 'Monday', 'First').
// ─────────────────────────────────────────────────────────────────────────────

export type RecurrencePatternType =
  | 'daily'
  | 'weekly'
  | 'absoluteMonthly'
  | 'relativeMonthly'
  | 'absoluteYearly'
  | 'relativeYearly';

export type WeekIndex = 'first' | 'second' | 'third' | 'fourth' | 'last';

export type DayOfWeek =
  | 'sunday' | 'monday' | 'tuesday' | 'wednesday' | 'thursday' | 'friday' | 'saturday';

export type RecurrenceRangeType = 'endDate' | 'noEnd' | 'numbered';

export interface RecurrenceInput {
  /** How often the event repeats. */
  pattern: RecurrencePatternType;
  /** Units between occurrences (e.g. every 3 weeks). Default 1. */
  interval?: number;
  /** Days the event falls on. Required for weekly, relativeMonthly, relativeYearly. */
  daysOfWeek?: DayOfWeek[];
  /** Day of the month (1-31). Required for absoluteMonthly and absoluteYearly. */
  dayOfMonth?: number;
  /** Month (1-12). Required for absoluteYearly and relativeYearly. */
  month?: number;
  /** Which occurrence of daysOfWeek in the month (e.g. 'last' Friday). Used by relative patterns. */
  index?: WeekIndex;
  /** First day of the week for weekly patterns (default 'sunday'). */
  firstDayOfWeek?: DayOfWeek;
  /** When the series stops. */
  range: {
    type: RecurrenceRangeType;
    /** Series start date (YYYY-MM-DD). Defaults to the event's start date. */
    startDate?: string;
    /** Last date of the series (YYYY-MM-DD). Required when type is 'endDate'. */
    endDate?: string;
    /** Total number of occurrences. Required when type is 'numbered'. */
    numberOfOccurrences?: number;
    /** Time zone the recurrence runs in. Defaults to the event time zone. */
    timeZone?: string;
  };
}

const capitalise = (s: string): string => s.charAt(0).toUpperCase() + s.slice(1);

/**
 * Validate a RecurrenceInput and translate it to the OWA `Recurrence` payload.
 * `eventStart` (the event's ISO start) seeds the range start date when omitted.
 */
export function buildRecurrence(
  input: RecurrenceInput,
  eventStart: string,
  defaultTimeZone: string,
): Record<string, unknown> {
  const pattern: Record<string, unknown> = {
    Type: capitalise(input.pattern),
    Interval: input.interval ?? 1,
  };

  const needsDays = ['weekly', 'relativeMonthly', 'relativeYearly'].includes(input.pattern);
  if (needsDays && !input.daysOfWeek?.length) {
    throw new Error(`Recurrence pattern '${input.pattern}' requires daysOfWeek.`);
  }
  if ((input.pattern === 'absoluteMonthly' || input.pattern === 'absoluteYearly') && input.dayOfMonth == null) {
    throw new Error(`Recurrence pattern '${input.pattern}' requires dayOfMonth.`);
  }
  if ((input.pattern === 'absoluteYearly' || input.pattern === 'relativeYearly') && input.month == null) {
    throw new Error(`Recurrence pattern '${input.pattern}' requires month (1-12).`);
  }

  if (input.daysOfWeek?.length) pattern.DaysOfWeek = input.daysOfWeek.map(capitalise);
  if (input.dayOfMonth != null) pattern.DayOfMonth = input.dayOfMonth;
  if (input.month != null) pattern.Month = input.month;
  if (input.index) pattern.Index = capitalise(input.index);
  if (input.pattern === 'weekly') pattern.FirstDayOfWeek = capitalise(input.firstDayOfWeek ?? 'sunday');

  const r = input.range;
  if (r.type === 'endDate' && !r.endDate) {
    throw new Error("Recurrence range type 'endDate' requires endDate (YYYY-MM-DD).");
  }
  if (r.type === 'numbered' && r.numberOfOccurrences == null) {
    throw new Error("Recurrence range type 'numbered' requires numberOfOccurrences.");
  }

  const startDate = r.startDate ?? eventStart.slice(0, 10);
  if (!startDate) {
    throw new Error('Recurrence requires a range startDate (or an event start to derive it from).');
  }

  const range: Record<string, unknown> = {
    Type: capitalise(r.type),
    StartDate: startDate,
    RecurrenceTimeZone: r.timeZone ?? defaultTimeZone,
  };
  if (r.endDate) range.EndDate = r.endDate;
  if (r.numberOfOccurrences != null) range.NumberOfOccurrences = r.numberOfOccurrences;

  return { Pattern: pattern, Range: range };
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

const WEEK_MS = 7 * 24 * 60 * 60 * 1000;

export async function listEvents(opts: ListEventsOptions = {}): Promise<CalendarEvent[]> {
  // calendarView (rather than /events) scopes results to [startDateTime, endDateTime]
  // and expands recurring series into individual instances within the window.
  const path = opts.calendarId ? `/calendars/${opts.calendarId}/calendarview` : '/calendarview';

  // Anchor the window on the start when given, so a future start without an end
  // produces a forward week rather than an empty/backwards window ending "now".
  const start = opts.startDateTime ? new Date(opts.startDateTime) : new Date();
  const startISO = opts.startDateTime ?? start.toISOString();
  const endISO = opts.endDateTime ?? new Date(start.getTime() + WEEK_MS).toISOString();

  const params: Record<string, string> = {
    startDateTime: startISO,
    endDateTime: endISO,
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

export type ShowAs = 'free' | 'tentative' | 'busy' | 'oof' | 'workingElsewhere';

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
  /** Make the event a repeating series. See RecurrenceInput. */
  recurrence?: RecurrenceInput;
  /**
   * Minutes before the start to fire the reminder. Note: Graph/OWA support only
   * ONE reminder per event, so a single value is all that can be set (you cannot
   * have, say, both a one-month and a one-week reminder on the same event).
   */
  reminderMinutesBeforeStart?: number;
  /** Whether the reminder is enabled. */
  isReminderOn?: boolean;
  /** Free/busy status to show for the event. */
  showAs?: ShowAs;
  /** Colour categories (labels) to tag the event with. */
  categories?: string[];
  /** Mark the event private (maps to Sensitivity 'Private'); otherwise 'Normal'. */
  isPrivate?: boolean;
}

/** Shared builder for the create/update payload fields that both endpoints accept. */
function buildEventExtras(opts: Partial<CreateEventOptions>, tz: string, eventStart?: string): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  if (opts.recurrence) out.Recurrence = buildRecurrence(opts.recurrence, eventStart ?? opts.start ?? '', tz);
  if (opts.reminderMinutesBeforeStart != null) out.ReminderMinutesBeforeStart = opts.reminderMinutesBeforeStart;
  if (opts.isReminderOn != null) out.IsReminderOn = opts.isReminderOn;
  if (opts.showAs) out.ShowAs = capitalise(opts.showAs);
  if (opts.categories) out.Categories = opts.categories;
  if (opts.isPrivate != null) out.Sensitivity = opts.isPrivate ? 'Private' : 'Normal';
  return out;
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
    ...buildEventExtras(opts, tz, opts.start),
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
  Object.assign(body, buildEventExtras(updates, tz, updates.start));

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
