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
  type CalendarEvent,
  type EventResponse,
} from '../api/calendar.js';

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
    },
    async ({ id, ...updates }) => {
      const event = await updateEvent(id, {
        subject: updates.subject,
        start: updates.start,
        end: updates.end,
        timeZone: updates.time_zone,
        location: updates.location,
        body: updates.body,
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
}
