/**
 * Calendar MCP tools.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import {
  listEvents,
  getEvent,
  createEvent,
  updateEvent,
  deleteEvent,
  respondToEvent,
  searchEvents,
  listCalendars,
  getSchedule,
  findMeetingTimes,
  cancelEvent,
  forwardEvent,
  type CalendarEvent,
  type EventResponse,
  type ScheduleInformation,
  type RecurrenceInput,
} from '../api/calendar.js';

// ─────────────────────────────────────────────────────────────────────────────
// Shared schemas
// ─────────────────────────────────────────────────────────────────────────────

const WEEKDAYS = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'] as const;

const recurrenceSchema = z.object({
  pattern: z.enum(['daily', 'weekly', 'absoluteMonthly', 'relativeMonthly', 'absoluteYearly', 'relativeYearly'])
    .describe('Repeat frequency. weekly needs days_of_week; absoluteMonthly/absoluteYearly need day_of_month; relativeMonthly/relativeYearly need days_of_week (+index); yearly patterns need month.'),
  interval: z.number().int().min(1).optional().describe('Units between occurrences, e.g. 3 = every 3 weeks/months (default 1)'),
  days_of_week: z.array(z.enum(WEEKDAYS)).optional().describe('Days the event falls on (e.g. ["monday"]). Required for weekly/relativeMonthly/relativeYearly'),
  day_of_month: z.number().int().min(1).max(31).optional().describe('Day of month (1-31). Required for absoluteMonthly/absoluteYearly'),
  month: z.number().int().min(1).max(12).optional().describe('Month 1-12. Required for absoluteYearly/relativeYearly'),
  index: z.enum(['first', 'second', 'third', 'fourth', 'last']).optional().describe('Which occurrence in the month, e.g. "last" Friday. Used by relative patterns'),
  first_day_of_week: z.enum(WEEKDAYS).optional().describe('First day of week for weekly patterns (default sunday)'),
  range: z.object({
    type: z.enum(['endDate', 'numbered', 'noEnd']).describe('How the series ends'),
    start_date: z.string().optional().describe('Series start date YYYY-MM-DD (defaults to the event start date)'),
    end_date: z.string().optional().describe('Last date YYYY-MM-DD. Required when type is endDate'),
    count: z.number().int().min(1).optional().describe('Total number of occurrences. Required when type is numbered'),
    time_zone: z.string().optional().describe('Recurrence time zone (defaults to the event time zone)'),
  }).describe('When the recurring series stops'),
}).describe('Make the event a repeating series (maps to Microsoft Graph patternedRecurrence).');

type RecurrenceArg = z.infer<typeof recurrenceSchema>;

function toRecurrenceInput(r: RecurrenceArg): RecurrenceInput {
  return {
    pattern: r.pattern,
    interval: r.interval,
    daysOfWeek: r.days_of_week,
    dayOfMonth: r.day_of_month,
    month: r.month,
    index: r.index,
    firstDayOfWeek: r.first_day_of_week,
    range: {
      type: r.range.type,
      startDate: r.range.start_date,
      endDate: r.range.end_date,
      numberOfOccurrences: r.range.count,
      timeZone: r.range.time_zone,
    },
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Formatters
// ─────────────────────────────────────────────────────────────────────────────

function formatEvent(e: CalendarEvent, full = false): string {
  const start = e.Start?.DateTime ?? 'Unknown';
  const end = e.End?.DateTime ?? 'Unknown';
  const tz = e.Start?.TimeZone ?? '';
  const location = e.Location?.DisplayName ?? '';
  const organizer = e.Organizer?.EmailAddress?.Address ?? '';
  const attendeeCount = e.Attendees?.length ?? 0;

  const lines = [
    `ID: ${e.Id}`,
    `Subject: ${e.Subject}`,
    `Start: ${start} ${tz}`,
    `End: ${end}`,
    location ? `Location: ${location}` : '',
    organizer ? `Organizer: ${organizer}` : '',
    attendeeCount > 0 ? `Attendees: ${attendeeCount}` : '',
    e.IsOnlineMeeting ? `Online Meeting: ${e.OnlineMeetingUrl ?? 'Yes'}` : '',
    e.Recurrence ? 'Recurring: yes' : '',
    e.IsReminderOn ? `Reminder: ${e.ReminderMinutesBeforeStart ?? 0} min before` : '',
    e.ShowAs ? `Show as: ${e.ShowAs}` : '',
    e.Categories?.length ? `Categories: ${e.Categories.join(', ')}` : '',
    e.ResponseStatus ? `My Response: ${e.ResponseStatus.Response}` : '',
    e.WebLink ? `Web URL: ${e.WebLink}` : '',
  ].filter(Boolean);

  if (full) {
    if (e.Attendees?.length) {
      lines.push('', 'Attendees:');
      lines.push(...e.Attendees.map(a =>
        `  - ${a.EmailAddress.Name ?? a.EmailAddress.Address} (${a.Type ?? 'Required'}) — ${a.Status?.Response ?? 'None'}`,
      ));
    }
    if (e.Body) {
      lines.push('', `Body (${e.Body.ContentType}):`, e.Body.Content.slice(0, 3000));
    }
  } else if (e.BodyPreview) {
    lines.push(`Preview: ${e.BodyPreview.slice(0, 200)}`);
  }

  return lines.join('\n');
}

function formatEventList(events: CalendarEvent[]): string {
  if (events.length === 0) return 'No events found.';
  return events.map((e, i) => `--- Event ${i + 1} ---\n${formatEvent(e)}`).join('\n\n');
}

function formatSchedule(s: ScheduleInformation): string {
  const lines = [`Person: ${s.ScheduleId}`];
  const busy = (s.ScheduleItems ?? []).filter(i => i.Status && i.Status !== 'Free');
  if (busy.length === 0) {
    lines.push('No busy blocks in the window (free).');
  } else {
    lines.push('Busy blocks:');
    lines.push(...busy.map(i => `  - ${i.Status}: ${i.Start?.DateTime ?? '?'} to ${i.End?.DateTime ?? '?'}${i.Subject ? ` (${i.Subject})` : ''}`));
  }
  const wh = s.WorkingHours;
  if (wh?.StartTime && wh.EndTime) {
    lines.push(`Working hours: ${wh.StartTime} to ${wh.EndTime}${wh.TimeZone?.Name ? ` ${wh.TimeZone.Name}` : ''}`);
  }
  return lines.join('\n');
}

// ─────────────────────────────────────────────────────────────────────────────
// Tools
// ─────────────────────────────────────────────────────────────────────────────

export function registerCalendarTools(server: McpServer): void {
  // ── outlook_list_events ──────────────────────────────────────────────────
  server.tool(
    'outlook_list_events',
    'List calendar events within a date range.',
    {
      start_date: z.string().optional().describe('ISO 8601 start datetime (defaults to now)'),
      end_date: z.string().optional().describe('ISO 8601 end datetime (defaults to 7 days from now)'),
      top: z.number().int().min(1).max(100).optional().describe('Max events to return (default 50)'),
      calendar_id: z.string().optional().describe('Specific calendar ID (defaults to primary calendar)'),
    },
    async ({ start_date, end_date, top, calendar_id }) => {
      const events = await listEvents({
        startDateTime: start_date,
        endDateTime: end_date,
        top: top ?? 50,
        calendarId: calendar_id,
      });
      return { content: [{ type: 'text', text: formatEventList(events) }] };
    },
  );

  // ── outlook_get_event ────────────────────────────────────────────────────
  server.tool(
    'outlook_get_event',
    'Get full details of a calendar event including attendees and body.',
    {
      id: z.string().describe('Event ID'),
    },
    async ({ id }) => {
      const event = await getEvent(id);
      return { content: [{ type: 'text', text: formatEvent(event, true) }] };
    },
  );

  // ── outlook_create_event ─────────────────────────────────────────────────
  server.tool(
    'outlook_create_event',
    'Create a new calendar event or meeting.',
    {
      subject: z.string().describe('Event title/subject'),
      start: z.string().describe('Start datetime in ISO 8601 format (e.g. 2026-06-10T14:00:00)'),
      end: z.string().describe('End datetime in ISO 8601 format'),
      time_zone: z.string().optional().describe('Timezone name (e.g. Europe/London, UTC). Defaults to UTC'),
      location: z.string().optional().describe('Location or meeting room'),
      body: z.string().optional().describe('Event description/body'),
      attendees: z.array(z.object({
        email: z.string().email(),
        name: z.string().optional(),
        type: z.enum(['Required', 'Optional']).optional(),
      })).optional().describe('List of attendees'),
      is_online_meeting: z.boolean().optional().describe('Add a Teams/online meeting link (default: false)'),
      is_all_day: z.boolean().optional().describe('All-day event (default: false)'),
      importance: z.enum(['Low', 'Normal', 'High']).optional(),
      recurrence: recurrenceSchema.optional(),
      reminder_minutes_before_start: z.number().int().min(0).optional().describe('Minutes before start to fire the reminder. Graph allows only ONE reminder per event'),
      is_reminder_on: z.boolean().optional().describe('Whether the reminder is enabled'),
      show_as: z.enum(['free', 'tentative', 'busy', 'oof', 'workingElsewhere']).optional().describe('Free/busy status shown for the event'),
      categories: z.array(z.string()).optional().describe('Colour category labels to tag the event with'),
      is_private: z.boolean().optional().describe('Mark the event private (sensitivity)'),
    },
    async (params) => {
      const event = await createEvent({
        subject: params.subject,
        start: params.start,
        end: params.end,
        timeZone: params.time_zone,
        location: params.location,
        body: params.body,
        attendees: params.attendees,
        isOnlineMeeting: params.is_online_meeting,
        isAllDay: params.is_all_day,
        importance: params.importance,
        recurrence: params.recurrence ? toRecurrenceInput(params.recurrence) : undefined,
        reminderMinutesBeforeStart: params.reminder_minutes_before_start,
        isReminderOn: params.is_reminder_on,
        showAs: params.show_as,
        categories: params.categories,
        isPrivate: params.is_private,
      });
      return {
        content: [{
          type: 'text',
          text: `Event created.\nID: ${event.Id}\nSubject: ${event.Subject}\nStart: ${event.Start.DateTime}`,
        }],
      };
    },
  );

  // ── outlook_update_event ─────────────────────────────────────────────────
  server.tool(
    'outlook_update_event',
    'Update an existing calendar event.',
    {
      id: z.string().describe('Event ID to update'),
      subject: z.string().optional().describe('New subject'),
      start: z.string().optional().describe('New start datetime (ISO 8601)'),
      end: z.string().optional().describe('New end datetime (ISO 8601)'),
      time_zone: z.string().optional(),
      location: z.string().optional(),
      body: z.string().optional(),
      recurrence: recurrenceSchema.optional().describe('Turn the event into a repeating series. Provide range.start_date if the event start is not also being set'),
      reminder_minutes_before_start: z.number().int().min(0).optional().describe('Minutes before start to fire the reminder. Graph allows only ONE reminder per event'),
      is_reminder_on: z.boolean().optional().describe('Whether the reminder is enabled'),
      show_as: z.enum(['free', 'tentative', 'busy', 'oof', 'workingElsewhere']).optional().describe('Free/busy status shown for the event'),
      categories: z.array(z.string()).optional().describe('Colour category labels to tag the event with'),
      is_private: z.boolean().optional().describe('Mark the event private (sensitivity)'),
    },
    async ({ id, ...updates }) => {
      const event = await updateEvent(id, {
        subject: updates.subject,
        start: updates.start,
        end: updates.end,
        timeZone: updates.time_zone,
        location: updates.location,
        body: updates.body,
        recurrence: updates.recurrence ? toRecurrenceInput(updates.recurrence) : undefined,
        reminderMinutesBeforeStart: updates.reminder_minutes_before_start,
        isReminderOn: updates.is_reminder_on,
        showAs: updates.show_as,
        categories: updates.categories,
        isPrivate: updates.is_private,
      });
      return {
        content: [{
          type: 'text',
          text: `Event updated.\nID: ${event.Id}\nSubject: ${event.Subject}`,
        }],
      };
    },
  );

  // ── outlook_delete_event ─────────────────────────────────────────────────
  server.tool(
    'outlook_delete_event',
    'Delete a calendar event.',
    {
      id: z.string().describe('Event ID to delete'),
    },
    async ({ id }) => {
      await deleteEvent(id);
      return { content: [{ type: 'text', text: 'Event deleted.' }] };
    },
  );

  // ── outlook_respond_to_event ─────────────────────────────────────────────
  server.tool(
    'outlook_respond_to_event',
    'Accept, decline, or tentatively accept a meeting invitation.',
    {
      id: z.string().describe('Event ID'),
      response: z.enum(['accept', 'decline', 'tentativelyAccept']).describe('Your response'),
      comment: z.string().optional().describe('Optional response message'),
    },
    async ({ id, response, comment }) => {
      await respondToEvent(id, response as EventResponse, comment);
      return { content: [{ type: 'text', text: `Meeting response sent: ${response}.` }] };
    },
  );

  // ── outlook_search_events ────────────────────────────────────────────────
  server.tool(
    'outlook_search_events',
    'Search calendar events by keyword.',
    {
      query: z.string().describe('Search keyword'),
      from_date: z.string().optional().describe('Search from this date (ISO 8601)'),
      to_date: z.string().optional().describe('Search until this date (ISO 8601)'),
    },
    async ({ query, from_date, to_date }) => {
      const events = await searchEvents(query, from_date, to_date);
      return { content: [{ type: 'text', text: formatEventList(events) }] };
    },
  );

  // ── outlook_list_calendars ───────────────────────────────────────────────
  server.tool(
    'outlook_list_calendars',
    'List all calendars in the account.',
    {},
    async () => {
      const calendars = await listCalendars();
      const text = calendars
        .map(c => `${c.Name} (ID: ${c.Id})${c.IsDefaultCalendar ? ' [Default]' : ''}${c.CanEdit ? ' [Editable]' : ''}`)
        .join('\n');
      return { content: [{ type: 'text', text: text || 'No calendars found.' }] };
    },
  );

  // ── outlook_get_schedule ─────────────────────────────────────────────────
  server.tool(
    'outlook_get_schedule',
    'Get the free/busy schedule for one or more people over a time window: their busy blocks (Busy, Tentative, OutOfOffice, WorkingElsewhere) and working hours. Use this to see when people are occupied. For a quick "are they free or out of office right now" summary use outlook_get_availability instead.',
    {
      emails: z.array(z.string().email()).min(1).max(50).describe('Email addresses to check'),
      start: z.string().optional().describe('ISO 8601 window start (defaults to now)'),
      end: z.string().optional().describe('ISO 8601 window end (defaults to 24 hours from start)'),
      interval_minutes: z.number().int().min(5).max(1440).optional().describe('Free/busy slot size in minutes (default 30)'),
    },
    async ({ emails, start, end, interval_minutes }) => {
      const startDate = start ? new Date(start) : new Date();
      const endDate = end ? new Date(end) : new Date(startDate.getTime() + 24 * 60 * 60 * 1000);
      const schedules = await getSchedule(emails, startDate, endDate, interval_minutes ?? 30);
      const text = schedules.length ? schedules.map(formatSchedule).join('\n\n') : 'No schedule information returned.';
      return { content: [{ type: 'text', text }] };
    },
  );

  // ── outlook_find_meeting_times ───────────────────────────────────────────
  server.tool(
    'outlook_find_meeting_times',
    'Suggest meeting time slots that work for a set of attendees, based on their free/busy. Returns ranked candidate slots with a confidence score and each attendee\'s availability. Useful for scheduling a meeting across several people.',
    {
      attendees: z.array(z.string().email()).min(1).max(20).describe('Attendee email addresses'),
      duration_minutes: z.number().int().min(15).max(480).optional().describe('Meeting length in minutes (default 30)'),
      start: z.string().optional().describe('ISO 8601 earliest start to consider (defaults to now)'),
      end: z.string().optional().describe('ISO 8601 latest end to consider (defaults to 5 days out)'),
      max_candidates: z.number().int().min(1).max(20).optional().describe('Max suggestions to return (default 5)'),
      time_zone: z.string().optional().describe('Time zone for the window (default UTC)'),
    },
    async ({ attendees, duration_minutes, start, end, max_candidates, time_zone }) => {
      const { suggestions, emptyReason } = await findMeetingTimes({
        attendees,
        durationMinutes: duration_minutes,
        start,
        end,
        maxCandidates: max_candidates,
        timeZone: time_zone,
      });
      if (suggestions.length === 0) {
        return { content: [{ type: 'text', text: `No meeting times found.${emptyReason ? ` Reason: ${emptyReason}` : ''}` }] };
      }
      const text = suggestions.map((s, i) => {
        const slot = s.MeetingTimeSlot;
        const who = (s.AttendeeAvailability ?? []).map(a => `${a.Attendee.EmailAddress.Address}: ${a.Availability}`).join(', ');
        return `--- Suggestion ${i + 1} (confidence ${s.Confidence ?? '?'}%) ---\n${slot?.Start.DateTime} to ${slot?.End.DateTime} ${slot?.Start.TimeZone ?? ''}\n${who}`;
      }).join('\n\n');
      return { content: [{ type: 'text', text }] };
    },
  );

  // ── outlook_cancel_event ─────────────────────────────────────────────────
  server.tool(
    'outlook_cancel_event',
    'Cancel an event that you organize, sending a cancellation to all attendees. This is different from outlook_delete_event (which just removes it from your calendar). Confirm with the user before calling.',
    {
      id: z.string().describe('Event ID to cancel'),
      comment: z.string().optional().describe('Optional cancellation message to attendees'),
    },
    async ({ id, comment }) => {
      await cancelEvent(id, comment);
      return { content: [{ type: 'text', text: 'Event cancelled and attendees notified.' }] };
    },
  );

  // ── outlook_forward_event ────────────────────────────────────────────────
  server.tool(
    'outlook_forward_event',
    'Forward a meeting invitation to additional people, effectively inviting them.',
    {
      id: z.string().describe('Event ID to forward'),
      to: z.array(z.string().email()).min(1).describe('People to forward the invite to'),
      comment: z.string().optional().describe('Optional message to include'),
    },
    async ({ id, to, comment }) => {
      await forwardEvent(id, to, comment);
      return { content: [{ type: 'text', text: `Invite forwarded to ${to.join(', ')}.` }] };
    },
  );
}
