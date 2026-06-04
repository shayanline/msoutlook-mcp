/**
 * Mailbox settings API: read settings and manage your own automatic replies
 * (out of office), via the OWA REST endpoint the session token can reach.
 */

import { owaGet, owaPatch } from './client.js';

export type AutomaticRepliesStatus = 'Disabled' | 'AlwaysEnabled' | 'Scheduled';
export type ExternalAudience = 'None' | 'ContactsOnly' | 'All';

interface OwaDateTime { DateTime: string; TimeZone: string }

export interface AutomaticRepliesSetting {
  Status?: AutomaticRepliesStatus;
  ExternalAudience?: ExternalAudience;
  InternalReplyMessage?: string;
  ExternalReplyMessage?: string;
  ScheduledStartDateTime?: OwaDateTime;
  ScheduledEndDateTime?: OwaDateTime;
}

export interface MailboxSettings {
  TimeZone?: string;
  DateFormat?: string;
  TimeFormat?: string;
  AutomaticRepliesSetting?: AutomaticRepliesSetting;
}

export async function getMailboxSettings(): Promise<MailboxSettings> {
  return owaGet<MailboxSettings>('/mailboxsettings');
}

export interface SetAutomaticRepliesOptions {
  status: AutomaticRepliesStatus;
  internalMessage?: string;
  externalMessage?: string;
  externalAudience?: ExternalAudience;
  /** ISO datetime; required when status is Scheduled. */
  start?: string;
  end?: string;
  timeZone?: string;
}

/**
 * Turn automatic replies on or off for the signed-in user. When status is
 * Scheduled, start and end define the window. The external message defaults to
 * the internal one when only the internal text is supplied.
 */
export async function setAutomaticReplies(opts: SetAutomaticRepliesOptions): Promise<MailboxSettings> {
  const tz = opts.timeZone ?? 'UTC';
  const setting: AutomaticRepliesSetting = {
    Status: opts.status,
    ExternalAudience: opts.externalAudience ?? 'All',
  };

  if (opts.status !== 'Disabled') {
    setting.InternalReplyMessage = opts.internalMessage ?? '';
    setting.ExternalReplyMessage = opts.externalMessage ?? opts.internalMessage ?? '';
  }
  if (opts.status === 'Scheduled') {
    if (opts.start) setting.ScheduledStartDateTime = { DateTime: opts.start, TimeZone: tz };
    if (opts.end) setting.ScheduledEndDateTime = { DateTime: opts.end, TimeZone: tz };
  }

  return owaPatch<MailboxSettings>('/mailboxsettings', { AutomaticRepliesSetting: setting });
}
