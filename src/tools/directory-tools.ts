/**
 * Directory (org hierarchy, profiles, photos) and mailbox-settings MCP tools.
 */

import { writeFile } from 'node:fs/promises';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import {
  getUserProfile,
  getManager,
  getDirectReports,
  getUserPhoto,
  type DirectoryUser,
} from '../api/directory.js';
import {
  getMailboxSettings,
  setAutomaticReplies,
  type AutomaticRepliesSetting,
} from '../api/mailbox.js';

function formatUser(u: DirectoryUser): string {
  return [
    `Name: ${u.displayName ?? '(unknown)'}`,
    u.mail ? `Email: ${u.mail}` : '',
    u.jobTitle ? `Title: ${u.jobTitle}` : '',
    u.department ? `Department: ${u.department}` : '',
    u.officeLocation ? `Office: ${u.officeLocation}` : '',
    u.mobilePhone ? `Mobile: ${u.mobilePhone}` : '',
    u.businessPhones?.length ? `Phone: ${u.businessPhones.join(', ')}` : '',
    [u.city, u.country].filter(Boolean).length ? `Location: ${[u.city, u.country].filter(Boolean).join(', ')}` : '',
  ].filter(Boolean).join('\n');
}

function formatAutomaticReplies(ar?: AutomaticRepliesSetting): string {
  if (!ar) return 'Automatic replies: unknown.';
  const lines = [`Automatic replies: ${ar.Status ?? 'Unknown'}`];
  if (ar.Status === 'Scheduled') {
    lines.push(`Window: ${ar.ScheduledStartDateTime?.DateTime ?? '?'} to ${ar.ScheduledEndDateTime?.DateTime ?? '?'}`);
  }
  if (ar.Status && ar.Status !== 'Disabled') {
    lines.push(`External audience: ${ar.ExternalAudience ?? 'All'}`);
    const internal = ar.InternalReplyMessage?.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
    if (internal) lines.push(`Internal message: ${internal}`);
    const external = ar.ExternalReplyMessage?.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
    if (external) lines.push(`External message: ${external}`);
  }
  return lines.join('\n');
}

export function registerDirectoryTools(server: McpServer): void {
  // ── outlook_get_user_profile ─────────────────────────────────────────────
  server.tool(
    'outlook_get_user_profile',
    'Get a colleague\'s directory profile (job title, department, office, phone) by email. Omit the email to get your own profile.',
    {
      email: z.string().email().optional().describe('Email/UPN of the person (omit for yourself)'),
    },
    async ({ email }) => {
      const user = await getUserProfile(email);
      return { content: [{ type: 'text', text: formatUser(user) }] };
    },
  );

  // ── outlook_get_manager ──────────────────────────────────────────────────
  server.tool(
    'outlook_get_manager',
    'Get a person\'s manager from the org directory by email. Omit the email to get your own manager.',
    {
      email: z.string().email().optional().describe('Email/UPN of the person (omit for yourself)'),
    },
    async ({ email }) => {
      try {
        const manager = await getManager(email);
        return { content: [{ type: 'text', text: formatUser(manager) }] };
      } catch {
        return { content: [{ type: 'text', text: 'No manager found for this person (or not visible in the directory).' }] };
      }
    },
  );

  // ── outlook_get_direct_reports ───────────────────────────────────────────
  server.tool(
    'outlook_get_direct_reports',
    'List a person\'s direct reports from the org directory by email. Omit the email for your own direct reports.',
    {
      email: z.string().email().optional().describe('Email/UPN of the manager (omit for yourself)'),
    },
    async ({ email }) => {
      const reports = await getDirectReports(email);
      if (reports.length === 0) return { content: [{ type: 'text', text: 'No direct reports.' }] };
      const text = reports.map((r, i) => `--- Report ${i + 1} ---\n${formatUser(r)}`).join('\n\n');
      return { content: [{ type: 'text', text }] };
    },
  );

  // ── outlook_get_user_photo ───────────────────────────────────────────────
  server.tool(
    'outlook_get_user_photo',
    'Download a person\'s profile photo and save it to a local file path. Omit the email for your own photo.',
    {
      output_path: z.string().describe('Local file path to write the photo to (e.g. ./photo.jpg)'),
      email: z.string().email().optional().describe('Email/UPN of the person (omit for yourself)'),
    },
    async ({ output_path, email }) => {
      try {
        const { bytes } = await getUserPhoto(email);
        await writeFile(output_path, bytes);
        return { content: [{ type: 'text', text: `Saved photo (${Math.round(bytes.length / 1024)}KB) to ${output_path}.` }] };
      } catch {
        return { content: [{ type: 'text', text: 'No photo available for this person.' }] };
      }
    },
  );

  // ── outlook_get_automatic_replies ────────────────────────────────────────
  server.tool(
    'outlook_get_automatic_replies',
    'Get your own out-of-office / automatic reply settings (status, schedule window, and messages).',
    {},
    async () => {
      const settings = await getMailboxSettings();
      return { content: [{ type: 'text', text: formatAutomaticReplies(settings.AutomaticRepliesSetting) }] };
    },
  );

  // ── outlook_set_automatic_replies ────────────────────────────────────────
  server.tool(
    'outlook_set_automatic_replies',
    'Turn your own out-of-office / automatic replies on or off. This changes your live mailbox, so confirm the message and dates with the user first. Use status Scheduled with start and end for a date window, AlwaysEnabled for on until turned off, or Disabled to turn off.',
    {
      status: z.enum(['Disabled', 'AlwaysEnabled', 'Scheduled']).describe('Disabled = off, AlwaysEnabled = on indefinitely, Scheduled = on for a window'),
      internal_message: z.string().optional().describe('Reply sent to people inside the organisation (HTML allowed)'),
      external_message: z.string().optional().describe('Reply sent to external senders (defaults to the internal message)'),
      external_audience: z.enum(['None', 'ContactsOnly', 'All']).optional().describe('Who outside the org gets a reply (default All)'),
      start: z.string().optional().describe('ISO 8601 start datetime, required when status is Scheduled'),
      end: z.string().optional().describe('ISO 8601 end datetime, required when status is Scheduled'),
      time_zone: z.string().optional().describe('Time zone for the schedule window (default UTC)'),
    },
    async ({ status, internal_message, external_message, external_audience, start, end, time_zone }) => {
      if (status === 'Scheduled' && (!start || !end)) {
        return { content: [{ type: 'text', text: 'Scheduled automatic replies need both start and end datetimes.' }] };
      }
      const settings = await setAutomaticReplies({
        status,
        internalMessage: internal_message,
        externalMessage: external_message,
        externalAudience: external_audience,
        start,
        end,
        timeZone: time_zone,
      });
      return { content: [{ type: 'text', text: `Automatic replies updated.\n${formatAutomaticReplies(settings.AutomaticRepliesSetting)}` }] };
    },
  );
}
