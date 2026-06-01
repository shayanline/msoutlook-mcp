/**
 * Mail API — email CRUD, search, and folder operations.
 */

import { owaGet, owaPost, owaPatch, owaDelete } from './client.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

export interface EmailAddress {
  Name: string;
  Address: string;
}

export interface Recipient {
  EmailAddress: EmailAddress;
}

export interface Message {
  Id: string;
  Subject: string;
  BodyPreview?: string;
  Body?: { ContentType: string; Content: string };
  From?: Recipient;
  Sender?: Recipient;
  ToRecipients?: Recipient[];
  CcRecipients?: Recipient[];
  BccRecipients?: Recipient[];
  ReceivedDateTime?: string;
  SentDateTime?: string;
  IsRead?: boolean;
  HasAttachments?: boolean;
  Importance?: string;
  Flag?: { FlagStatus: string };
  ConversationId?: string;
  ParentFolderId?: string;
  WebLink?: string;
  Attachments?: Attachment[];
}

export interface Attachment {
  Id: string;
  Name: string;
  ContentType: string;
  Size: number;
  IsInline: boolean;
}

export interface MailFolder {
  Id: string;
  DisplayName: string;
  UnreadItemCount: number;
  TotalItemCount: number;
  ChildFolderCount: number;
}

export interface ODataResponse<T> {
  value: T[];
  '@odata.nextLink'?: string;
  '@odata.count'?: number;
}

// ─────────────────────────────────────────────────────────────────────────────
// List messages
// ─────────────────────────────────────────────────────────────────────────────

export interface ListMessagesOptions {
  folder?: string;
  top?: number;
  skip?: number;
  filter?: string;
  select?: string[];
  orderBy?: string;
  search?: string;
}

export async function listMessages(opts: ListMessagesOptions = {}): Promise<Message[]> {
  const folder = opts.folder ?? 'Inbox';
  const params: Record<string, string> = {
    '$top': String(opts.top ?? 20),
    '$select': (opts.select ?? [
      'Id', 'Subject', 'BodyPreview', 'From', 'ToRecipients', 'ReceivedDateTime',
      'IsRead', 'HasAttachments', 'Importance', 'Flag', 'ConversationId', 'WebLink',
    ]).join(','),
  };

  if (opts.skip) params['$skip'] = String(opts.skip);

  // $search cannot be combined with $orderby or $filter — the API returns
  // 400 SearchWithOrderBy. Search results are relevance-ranked, so we drop
  // ordering when searching.
  if (opts.search) {
    params['$search'] = `"${opts.search}"`;
  } else {
    params['$orderby'] = opts.orderBy ?? 'ReceivedDateTime desc';
    if (opts.filter) params['$filter'] = opts.filter;
  }

  const res = await owaGet<ODataResponse<Message>>(`/MailFolders/${folder}/messages`, params);
  return res.value;
}

// ─────────────────────────────────────────────────────────────────────────────
// Get message
// ─────────────────────────────────────────────────────────────────────────────

export async function getMessage(id: string, includeAttachments = false): Promise<Message> {
  const params = includeAttachments ? { '$expand': 'Attachments' } : undefined;
  return owaGet<Message>(`/messages/${id}`, params);
}

// ─────────────────────────────────────────────────────────────────────────────
// Send email
// ─────────────────────────────────────────────────────────────────────────────

export interface SendEmailOptions {
  to: string[];
  cc?: string[];
  bcc?: string[];
  subject: string;
  body: string;
  bodyType?: 'Text' | 'HTML';
  importance?: 'Low' | 'Normal' | 'High';
  saveToSentItems?: boolean;
}

export async function sendEmail(opts: SendEmailOptions): Promise<void> {
  const toRecipients = opts.to.map(addr => ({
    EmailAddress: { Address: addr },
  }));
  const ccRecipients = opts.cc?.map(addr => ({ EmailAddress: { Address: addr } }));
  const bccRecipients = opts.bcc?.map(addr => ({ EmailAddress: { Address: addr } }));

  await owaPost('/sendmail', {
    Message: {
      Subject: opts.subject,
      Body: { ContentType: opts.bodyType ?? 'Text', Content: opts.body },
      ToRecipients: toRecipients,
      ...(ccRecipients?.length ? { CcRecipients: ccRecipients } : {}),
      ...(bccRecipients?.length ? { BccRecipients: bccRecipients } : {}),
      Importance: opts.importance ?? 'Normal',
    },
    SaveToSentItems: opts.saveToSentItems !== false,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Create draft
// ─────────────────────────────────────────────────────────────────────────────

export async function createDraft(opts: Omit<SendEmailOptions, 'saveToSentItems'>): Promise<Message> {
  const toRecipients = opts.to.map(addr => ({ EmailAddress: { Address: addr } }));
  const ccRecipients = opts.cc?.map(addr => ({ EmailAddress: { Address: addr } }));
  const bccRecipients = opts.bcc?.map(addr => ({ EmailAddress: { Address: addr } }));

  return owaPost<Message>('/messages', {
    Subject: opts.subject,
    Body: { ContentType: opts.bodyType ?? 'Text', Content: opts.body },
    ToRecipients: toRecipients,
    ...(ccRecipients?.length ? { CcRecipients: ccRecipients } : {}),
    ...(bccRecipients?.length ? { BccRecipients: bccRecipients } : {}),
    Importance: opts.importance ?? 'Normal',
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Reply
// ─────────────────────────────────────────────────────────────────────────────

export async function replyToMessage(id: string, body: string, replyAll = false): Promise<void> {
  const action = replyAll ? 'replyall' : 'reply';
  await owaPost(`/messages/${id}/${action}`, {
    Comment: body,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Forward
// ─────────────────────────────────────────────────────────────────────────────

export async function forwardMessage(id: string, to: string[], comment?: string): Promise<void> {
  await owaPost(`/messages/${id}/forward`, {
    ToRecipients: to.map(addr => ({ EmailAddress: { Address: addr } })),
    Comment: comment ?? '',
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Mark read/unread
// ─────────────────────────────────────────────────────────────────────────────

export async function markMessageRead(id: string, isRead = true): Promise<void> {
  await owaPatch(`/messages/${id}`, { IsRead: isRead });
}

// ─────────────────────────────────────────────────────────────────────────────
// Flag
// ─────────────────────────────────────────────────────────────────────────────

export async function flagMessage(id: string, status: 'Flagged' | 'Complete' | 'NotFlagged' = 'Flagged'): Promise<void> {
  await owaPatch(`/messages/${id}`, { Flag: { FlagStatus: status } });
}

// ─────────────────────────────────────────────────────────────────────────────
// Move
// ─────────────────────────────────────────────────────────────────────────────

export async function moveMessage(id: string, destinationFolderId: string): Promise<Message> {
  return owaPost<Message>(`/messages/${id}/move`, {
    DestinationId: destinationFolderId,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Delete
// ─────────────────────────────────────────────────────────────────────────────

export async function deleteMessage(id: string): Promise<void> {
  await owaDelete(`/messages/${id}`);
}

// ─────────────────────────────────────────────────────────────────────────────
// Search
// ─────────────────────────────────────────────────────────────────────────────

export interface SearchMessagesOptions {
  query: string;
  top?: number;
  folder?: string;
  /** Inclusive lower bound on received date (ISO date or datetime). */
  startDate?: string;
  /** Inclusive upper bound on received date (ISO date or datetime). */
  endDate?: string;
  /** Opaque page token from a previous result's nextSkipToken. */
  skipToken?: string;
}

export interface SearchMessagesResult {
  messages: Message[];
  /** Pass this back as skipToken to fetch the next page; absent when there are no more results. */
  nextSkipToken?: string;
}

/** Reduce an ISO date or datetime to the YYYY-MM-DD form KQL expects. */
function toKqlDate(value: string): string {
  return value.slice(0, 10);
}

/**
 * Search messages with optional date range and pagination.
 *
 * The OWA REST API rejects $skip, $filter and $orderby alongside $search, so:
 * - date range is expressed as KQL constraints inside the search string
 *   (received>=START AND received<=END), and
 * - pagination uses the $skiptoken returned in @odata.nextLink.
 */
export async function searchMessages(opts: SearchMessagesOptions): Promise<SearchMessagesResult> {
  const folder = opts.folder ?? 'Inbox';

  let kql = opts.query.trim();
  if (opts.startDate) kql += ` AND received>=${toKqlDate(opts.startDate)}`;
  if (opts.endDate) kql += ` AND received<=${toKqlDate(opts.endDate)}`;

  const params: Record<string, string> = {
    '$search': `"${kql}"`,
    '$top': String(opts.top ?? 20),
    '$select': 'Id,Subject,BodyPreview,From,ReceivedDateTime,IsRead,WebLink',
  };
  if (opts.skipToken) params['$skiptoken'] = opts.skipToken;

  const res = await owaGet<ODataResponse<Message>>(`/MailFolders/${folder}/messages`, params);

  let nextSkipToken: string | undefined;
  const nextLink = res['@odata.nextLink'];
  if (nextLink) {
    try {
      nextSkipToken = new URL(nextLink).searchParams.get('$skiptoken') ?? undefined;
    } catch {
      nextSkipToken = undefined;
    }
  }

  return { messages: res.value, nextSkipToken };
}

// ─────────────────────────────────────────────────────────────────────────────
// Folders
// ─────────────────────────────────────────────────────────────────────────────

export async function listFolders(): Promise<MailFolder[]> {
  const res = await owaGet<ODataResponse<MailFolder>>('/mailfolders', {
    '$top': '50',
    '$select': 'Id,DisplayName,UnreadItemCount,TotalItemCount,ChildFolderCount',
  });
  return res.value;
}

export async function getFolder(idOrName: string): Promise<MailFolder> {
  return owaGet<MailFolder>(`/mailfolders/${idOrName}`);
}

// ─────────────────────────────────────────────────────────────────────────────
// Unread count
// ─────────────────────────────────────────────────────────────────────────────

export async function getUnreadMessages(top = 10): Promise<Message[]> {
  return listMessages({
    filter: 'IsRead eq false',
    top,
    orderBy: 'ReceivedDateTime desc',
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Send draft
// ─────────────────────────────────────────────────────────────────────────────

export async function sendDraft(draftId: string): Promise<void> {
  await owaPost(`/messages/${draftId}/send`, {});
}
