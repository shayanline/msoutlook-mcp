/**
 * Availability API: is a person free or busy right now, and are they away.
 *
 * Built on two OWA REST endpoints that the Outlook session token can reach:
 * - /getmailtips         -> out-of-office / automatic reply status, and
 * - /calendar/getschedule -> free/busy view, working hours and time zone.
 *
 * Note on live Teams presence (the green/red/yellow dot): that comes from the
 * Presence.Read.All Graph scope, which the Outlook first-party app does not
 * grant, so it is not available here. Use the Teams MCP for the presence dot.
 * Free/busy from the calendar is the equivalent signal we can read from Outlook.
 */

import { owaPost } from './client.js';
import { getSchedule, type ScheduleInformation } from './calendar.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

export type FreeBusyStatus = 'Free' | 'Tentative' | 'Busy' | 'OutOfOffice' | 'WorkingElsewhere' | 'Unknown';

export interface OutOfOffice {
  isActive: boolean;
  message?: string;
  scheduledStart?: string;
  scheduledEnd?: string;
}

export interface WorkingHours {
  daysOfWeek?: string[];
  startTime?: string;
  endTime?: string;
  timeZone?: string;
}

export interface AvailabilityInfo {
  email: string;
  displayName?: string;
  /** Free/busy status at the current moment, from the calendar. */
  status?: FreeBusyStatus;
  /** ISO time at which the current status next changes, if known. */
  statusChangesAt?: string;
  workingHours?: WorkingHours;
  outOfOffice?: OutOfOffice;
  scheduleError?: string;
  outOfOfficeError?: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// OWA response shapes (PascalCase)
// ─────────────────────────────────────────────────────────────────────────────

interface OwaDateTime { DateTime: string; TimeZone: string }

interface MailTipsValue {
  EmailAddress?: { Name?: string; Address?: string };
  AutomaticReplies?: {
    Message?: string;
    ScheduledStartTime?: OwaDateTime;
    ScheduledEndTime?: OwaDateTime;
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Low level OWA calls
// ─────────────────────────────────────────────────────────────────────────────

/** Out-of-office / automatic reply mail tips for a batch of addresses. */
export async function getMailTips(emails: string[]): Promise<MailTipsValue[]> {
  const res = await owaPost<{ value: MailTipsValue[] }>('/getmailtips', {
    EmailAddresses: emails,
    MailTipsOptions: 'automaticReplies',
  });
  return res.value ?? [];
}

// ─────────────────────────────────────────────────────────────────────────────
// Combine into availability
// ─────────────────────────────────────────────────────────────────────────────

const VIEW_STATUS: Record<string, FreeBusyStatus> = {
  '0': 'Free',
  '1': 'Tentative',
  '2': 'Busy',
  '3': 'OutOfOffice',
  '4': 'WorkingElsewhere',
};

function shortError(err: unknown): string {
  const msg = err instanceof Error ? err.message : String(err);
  return msg.slice(0, 160);
}

function toOutOfOffice(ar: MailTipsValue['AutomaticReplies']): OutOfOffice {
  const message = ar?.Message?.trim();
  if (!message) return { isActive: false };
  return {
    isActive: true,
    message,
    scheduledStart: ar?.ScheduledStartTime?.DateTime,
    scheduledEnd: ar?.ScheduledEndTime?.DateTime,
  };
}

/**
 * Read the current free/busy status and, where it changes within the window,
 * the time of the next change. AvailabilityView is one digit per interval slot
 * starting at `start`, so index 0 is "now".
 */
function readSchedule(s: ScheduleInformation | undefined, start: Date, intervalMinutes: number): {
  status?: FreeBusyStatus;
  changesAt?: string;
  workingHours?: WorkingHours;
} {
  if (!s) return {};
  const view = s.AvailabilityView ?? '';
  const status = view ? (VIEW_STATUS[view[0]] ?? 'Unknown') : undefined;

  let changesAt: string | undefined;
  if (view) {
    const changeIndex = [...view].findIndex(c => c !== view[0]);
    if (changeIndex > 0) {
      changesAt = new Date(start.getTime() + changeIndex * intervalMinutes * 60_000).toISOString();
    }
  }

  const wh = s.WorkingHours;
  const workingHours = wh
    ? { daysOfWeek: wh.DaysOfWeek, startTime: wh.StartTime, endTime: wh.EndTime, timeZone: wh.TimeZone?.Name }
    : undefined;

  return { status, changesAt, workingHours };
}

/**
 * Get a combined availability view (free/busy now + out-of-office) for one or
 * more people by email. Always returns one entry per requested email, with
 * partial data and per-section error notes rather than failing the whole call.
 */
export async function getAvailability(emails: string[], windowHours = 8): Promise<AvailabilityInfo[]> {
  const interval = 30;
  const start = new Date();
  const end = new Date(start.getTime() + windowHours * 60 * 60 * 1000);

  const [tipsResult, scheduleResult] = await Promise.allSettled([
    getMailTips(emails),
    getSchedule(emails, start, end, interval),
  ]);

  const oofByEmail = new Map<string, OutOfOffice>();
  const nameByEmail = new Map<string, string>();
  let oofError: string | undefined;
  if (tipsResult.status === 'fulfilled') {
    for (const t of tipsResult.value) {
      const addr = t.EmailAddress?.Address?.toLowerCase();
      if (!addr) continue;
      oofByEmail.set(addr, toOutOfOffice(t.AutomaticReplies));
      if (t.EmailAddress?.Name) nameByEmail.set(addr, t.EmailAddress.Name);
    }
  } else {
    oofError = shortError(tipsResult.reason);
  }

  const scheduleByEmail = new Map<string, ScheduleInformation>();
  let scheduleError: string | undefined;
  if (scheduleResult.status === 'fulfilled') {
    for (const s of scheduleResult.value) scheduleByEmail.set(s.ScheduleId.toLowerCase(), s);
  } else {
    scheduleError = shortError(scheduleResult.reason);
  }

  return emails.map(email => {
    const key = email.toLowerCase();
    const { status, changesAt, workingHours } = readSchedule(scheduleByEmail.get(key), start, interval);

    const info: AvailabilityInfo = {
      email,
      displayName: nameByEmail.get(key) || undefined,
      status,
      statusChangesAt: changesAt,
      workingHours,
      outOfOffice: oofByEmail.get(key),
    };
    if (!status && scheduleError) info.scheduleError = scheduleError;
    if (!info.outOfOffice && oofError) info.outOfOfficeError = oofError;
    return info;
  });
}
