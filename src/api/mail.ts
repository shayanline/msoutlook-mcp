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
    '$orderby': opts.orderBy ?? 'ReceivedDateTime desc',
    '$select': (opts.select ?? [
      'Id', 'Subject', 'BodyPreview', 'From', 'ToRecipients', 'ReceivedDateTime',
      'IsRead', 'HasAttachments', 'Importance', 'Flag', 'ConversationId', 'WebLink',
    ]).join(','),
  };

  if (opts.skip) params['$skip'] = String(opts.skip);
  if (opts.filter) params['$filter'] = opts.filter;
  if (opts.search) params['$search'] = `"${opts.search}"`;

  const res = await owaGet<ODataResponse<Message>>(`/MailFolders/${folder}/messages`, params);
  return res.value;
}

// ─────────────────────────────────────────────────────────────────────────────
// Get message
// ─────────────────────────────────────────────────────────────────────────────

export async function getMessage(id: string, includeAttachments = false): Promise<Message> {
  const params: Record<string, string> = {
    '$expand': includeAttachments ? 'Attachments' : '',
  };
  if (!includeAttachments) delete params['$expand'];
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

export async function searchMessages(query: string, top = 20): Promise<Message[]> {
  return listMessages({
    search: query,
    top,
    select: ['Id', 'Subject', 'BodyPreview', 'From', 'ReceivedDateTime', 'IsRead', 'WebLink'],
  });
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
