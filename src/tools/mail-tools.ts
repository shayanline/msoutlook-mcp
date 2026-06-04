/**
 * Email MCP tools.
 */

import { writeFile } from 'node:fs/promises';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import {
  listMessages,
  getMessage,
  sendEmail,
  createDraft,
  replyToMessage,
  createReplyDraft,
  createForwardDraft,
  forwardMessage,
  markMessageRead,
  flagMessage,
  moveMessage,
  deleteMessage,
  searchMessages,
  listFolders,
  getUnreadMessages,
  sendDraft,
  updateDraft,
  fileToAttachment,
  listAttachments,
  getAttachmentContent,
  addAttachment,
  getConversation,
  setCategories,
  createFolder,
  renameFolder,
  deleteFolder,
  type Message,
  type Attachment,
} from '../api/mail.js';

// ─────────────────────────────────────────────────────────────────────────────
// Formatters
// ─────────────────────────────────────────────────────────────────────────────

function formatMessage(m: Message, full = false): string {
  const from = m.From?.EmailAddress
    ? `${m.From.EmailAddress.Name ?? ''} <${m.From.EmailAddress.Address}>`
    : 'Unknown';
  const to = m.ToRecipients?.map(r => r.EmailAddress.Address).join(', ') ?? '';
  const date = m.ReceivedDateTime ?? m.SentDateTime ?? '';

  const lines = [
    `ID: ${m.Id}`,
    `Subject: ${m.Subject}`,
    `From: ${from}`,
    `To: ${to}`,
    `Date: ${date}`,
    `Read: ${m.IsRead ? 'Yes' : 'No'}`,
    `Has Attachments: ${m.HasAttachments ? 'Yes' : 'No'}`,
    m.Flag?.FlagStatus !== 'NotFlagged' ? `Flag: ${m.Flag?.FlagStatus}` : '',
    m.WebLink ? `Web URL: ${m.WebLink}` : '',
  ].filter(Boolean);

  if (full && m.Body) {
    lines.push('', `Body (${m.Body.ContentType}):`, m.Body.Content.slice(0, 5000));
  } else if (m.BodyPreview) {
    lines.push(`Preview: ${m.BodyPreview}`);
  }

  return lines.join('\n');
}

function formatMessageList(messages: Message[]): string {
  if (messages.length === 0) return 'No messages found.';
  return messages.map((m, i) => `--- Message ${i + 1} ---\n${formatMessage(m)}`).join('\n\n');
}

/**
 * Run an async operation over many IDs in capped concurrent chunks.
 * One failure never aborts the rest: every result is captured.
 */
async function runBatch(
  ids: string[],
  op: (id: string) => Promise<unknown>,
  concurrency = 8,
): Promise<{ succeeded: number; failed: Array<{ id: string; error: string }> }> {
  let succeeded = 0;
  const failed: Array<{ id: string; error: string }> = [];
  for (let i = 0; i < ids.length; i += concurrency) {
    const chunk = ids.slice(i, i + concurrency);
    const results = await Promise.allSettled(chunk.map(op));
    results.forEach((r, j) => {
      if (r.status === 'fulfilled') succeeded += 1;
      else failed.push({ id: chunk[j], error: r.reason instanceof Error ? r.reason.message : String(r.reason) });
    });
  }
  return { succeeded, failed };
}

// ─────────────────────────────────────────────────────────────────────────────
// Tools
// ─────────────────────────────────────────────────────────────────────────────

export function registerMailTools(server: McpServer): void {
  // ── outlook_list_emails ──────────────────────────────────────────────────
  server.tool(
    'outlook_list_emails',
    'List emails from a mail folder. Defaults to Inbox, newest first.',
    {
      folder: z.string().optional().describe('Folder name or ID. Defaults to Inbox. Common values: Inbox, Drafts, SentItems, DeletedItems, Archive, JunkEmail'),
      top: z.number().int().min(1).max(100).optional().describe('Number of emails to return (default 20, max 100)'),
      unread_only: z.boolean().optional().describe('If true, return only unread emails'),
      mark_read: z.boolean().optional().describe('If true, mark fetched emails as read. Default: false'),
    },
    async ({ folder, top, unread_only, mark_read }) => {
      const messages = await listMessages({
        folder: folder ?? 'Inbox',
        top: top ?? 20,
        filter: unread_only ? 'IsRead eq false' : undefined,
      });

      if (mark_read) {
        await Promise.all(messages.map(m => markMessageRead(m.Id)));
      }

      return { content: [{ type: 'text', text: formatMessageList(messages) }] };
    },
  );

  // ── outlook_get_email ────────────────────────────────────────────────────
  server.tool(
    'outlook_get_email',
    'Read the full content of a specific email by ID.',
    {
      id: z.string().describe('The email message ID (from outlook_list_emails)'),
      include_attachments: z.boolean().optional().describe('Include attachment metadata (default: false)'),
      mark_read: z.boolean().optional().describe('Mark as read when fetching (default: false)'),
    },
    async ({ id, include_attachments, mark_read }) => {
      const message = await getMessage(id, include_attachments ?? false);

      if (mark_read && !message.IsRead) {
        await markMessageRead(id);
      }

      let text = formatMessage(message, true);

      if (include_attachments && message.Attachments?.length) {
        text += '\n\nAttachments:\n';
        text += message.Attachments.map(a => `- ${a.Name} (${a.ContentType}, ${Math.round(a.Size / 1024)}KB)`).join('\n');
      }

      return { content: [{ type: 'text', text }] };
    },
  );

  // ── outlook_get_unread ───────────────────────────────────────────────────
  server.tool(
    'outlook_get_unread',
    'Get a list of unread emails from the Inbox.',
    {
      top: z.number().int().min(1).max(50).optional().describe('Number to return (default 10)'),
    },
    async ({ top }) => {
      const messages = await getUnreadMessages(top ?? 10);
      const count = messages.length;
      const header = `${count} unread email${count !== 1 ? 's' : ''}:\n\n`;
      return { content: [{ type: 'text', text: header + formatMessageList(messages) }] };
    },
  );

  // ── outlook_send_email ───────────────────────────────────────────────────
  server.tool(
    'outlook_send_email',
    'Send an email immediately. Body format defaults to HTML (body_type HTML). Prefer the review first flow: unless the user has asked to send straight away, create the message with outlook_create_draft so they can review it, then send with outlook_send_draft once approved. Always confirm content with the user before calling this tool. Write structure with HTML (<br> for a line break, <br><br> for a paragraph gap, <ul><li>...</li></ul> for lists). Plain text is still accepted and its newlines are converted to <br> automatically, so a multi line message never arrives as one block.',
    {
      to: z.array(z.string().email()).describe('List of recipient email addresses'),
      cc: z.array(z.string().email()).optional().describe('CC recipients'),
      bcc: z.array(z.string().email()).optional().describe('BCC recipients'),
      subject: z.string().describe('Email subject'),
      body: z.string().describe('Email body. Prefer HTML markup (<br>, <div>, <ul><li>) for layout. Plain text is fine too: its newlines are auto converted to <br> so line breaks survive.'),
      body_type: z.enum(['Text', 'HTML']).optional().describe('Body format. Default: HTML, which is recommended so line breaks render. Only use Text to send a literal plain text body with no auto formatting.'),
      importance: z.enum(['Low', 'Normal', 'High']).optional().describe('Email importance (default: Normal)'),
      attachments: z.array(z.string()).optional().describe('Local file paths to attach. Each file is read from disk and attached.'),
    },
    async ({ to, cc, bcc, subject, body, body_type, importance, attachments }) => {
      const files = attachments?.length ? await Promise.all(attachments.map(fileToAttachment)) : undefined;
      await sendEmail({
        to,
        cc,
        bcc,
        subject,
        body,
        bodyType: body_type,
        importance,
        attachments: files,
      });
      const note = files?.length ? ` with ${files.length} attachment(s)` : '';
      return { content: [{ type: 'text', text: `Email sent to ${to.join(', ')}${note}.` }] };
    },
  );

  // ── outlook_create_draft ─────────────────────────────────────────────────
  server.tool(
    'outlook_create_draft',
    'Create a draft email without sending it. This is the preferred way to compose a new email: create the draft here so the user can review it in Outlook, then send it with outlook_send_draft once they approve, unless the user has asked to send straight away. Body format defaults to HTML (body_type HTML): use HTML for layout (<br>, <br><br>, <ul><li>). Plain text newlines are auto converted to <br> so the draft keeps its line breaks.',
    {
      to: z.array(z.string().email()).describe('Recipient email addresses'),
      cc: z.array(z.string().email()).optional().describe('CC recipients'),
      subject: z.string().describe('Email subject'),
      body: z.string().describe('Email body. Prefer HTML markup for layout. Plain text is fine too: its newlines are auto converted to <br>.'),
      body_type: z.enum(['Text', 'HTML']).optional().describe('Body format. Default: HTML. Only use Text for a literal plain text body with no auto formatting.'),
      attachments: z.array(z.string()).optional().describe('Local file paths to attach to the draft.'),
    },
    async ({ to, cc, subject, body, body_type, attachments }) => {
      const files = attachments?.length ? await Promise.all(attachments.map(fileToAttachment)) : undefined;
      const draft = await createDraft({ to, cc, subject, body, bodyType: body_type, attachments: files });
      return {
        content: [{
          type: 'text',
          text: `Draft created.\nID: ${draft.Id}\nSubject: ${draft.Subject}\nTo: ${to.join(', ')}`,
        }],
      };
    },
  );

  // ── outlook_send_draft ───────────────────────────────────────────────────
  server.tool(
    'outlook_send_draft',
    'Send a previously created draft email by its ID.',
    {
      id: z.string().describe('Draft message ID'),
    },
    async ({ id }) => {
      await sendDraft(id);
      return { content: [{ type: 'text', text: 'Draft sent successfully.' }] };
    },
  );

  // ── outlook_reply ────────────────────────────────────────────────────────
  server.tool(
    'outlook_reply',
    'Reply to an email, staying in the same thread and keeping all recipients when reply_all is true. This sends immediately, there is no separate reply draft, so confirm the content with the user before calling unless they have asked to send straight away. The reply is always rendered as HTML, so use HTML for layout (<br>, <br><br>, <ul><li>). Plain text is accepted and its newlines are auto converted to <br>, so a multi paragraph reply never collapses into one block.',
    {
      id: z.string().describe('Message ID to reply to'),
      body: z.string().describe('Reply body. Prefer HTML markup (<br>, <ul><li>) for layout. Plain text is fine too: its newlines are auto converted to <br>.'),
      reply_all: z.boolean().optional().describe('If true, reply to all recipients and keep every CC on the thread (default: false)'),
    },
    async ({ id, body, reply_all }) => {
      await replyToMessage(id, body, reply_all ?? false);
      return { content: [{ type: 'text', text: 'Reply sent.' }] };
    },
  );

  // ── outlook_create_reply_draft ───────────────────────────────────────────
  server.tool(
    'outlook_create_reply_draft',
    'Create a reply (or reply-all) as a DRAFT instead of sending it. This is the review-first way to reply: the draft is saved to Drafts with the recipients and quoted original prefilled and your text inserted above the quote, so the user can review or edit it in Outlook, then send it with outlook_send_draft once approved. Prefer this over outlook_reply unless the user has asked to send straight away. Body is rendered as HTML: use HTML for layout (<br>, <br><br>, <ul><li>); plain text newlines are auto converted to <br>.',
    {
      id: z.string().describe('Message ID to reply to'),
      body: z.string().describe('Reply body. Prefer HTML markup (<br>, <ul><li>) for layout. Plain text is fine too: its newlines are auto converted to <br>.'),
      reply_all: z.boolean().optional().describe('If true, reply to all recipients and keep every CC on the thread (default: false)'),
    },
    async ({ id, body, reply_all }) => {
      const draft = await createReplyDraft(id, body, reply_all ?? false);
      return {
        content: [{
          type: 'text',
          text: `Reply draft created and saved to Drafts.\nID: ${draft.Id}\nSubject: ${draft.Subject}\n\nReview or edit it in Outlook, then send with outlook_send_draft (id above).`,
        }],
      };
    },
  );

  // ── outlook_create_forward_draft ─────────────────────────────────────────
  server.tool(
    'outlook_create_forward_draft',
    'Create a forward as a DRAFT instead of sending it. The draft is saved to Drafts with the quoted original prefilled; recipients can be set here or added later in Outlook. Review or edit, then send with outlook_send_draft. Body is rendered as HTML: use HTML for layout; plain text newlines are auto converted to <br>.',
    {
      id: z.string().describe('Message ID to forward'),
      to: z.array(z.string().email()).optional().describe('Optional recipients to prefill on the forward draft'),
      comment: z.string().optional().describe('Optional message to add above the forwarded content. HTML preferred; plain text newlines are auto converted to <br>.'),
    },
    async ({ id, to, comment }) => {
      const draft = await createForwardDraft(id, comment ?? '', to);
      return {
        content: [{
          type: 'text',
          text: `Forward draft created and saved to Drafts.\nID: ${draft.Id}\nSubject: ${draft.Subject}\n\nReview, set or confirm recipients, then send with outlook_send_draft (id above).`,
        }],
      };
    },
  );

  // ── outlook_forward ──────────────────────────────────────────────────────
  server.tool(
    'outlook_forward',
    'Forward an email to one or more recipients.',
    {
      id: z.string().describe('Message ID to forward'),
      to: z.array(z.string().email()).describe('Forward to these addresses'),
      comment: z.string().optional().describe('Optional message to include with the forward. Rendered as HTML: use HTML for layout, or plain text whose newlines are auto converted to <br>.'),
    },
    async ({ id, to, comment }) => {
      await forwardMessage(id, to, comment);
      return { content: [{ type: 'text', text: `Forwarded to ${to.join(', ')}.` }] };
    },
  );

  // ── outlook_mark_read ────────────────────────────────────────────────────
  server.tool(
    'outlook_mark_read',
    'Mark an email as read or unread.',
    {
      id: z.string().describe('Message ID'),
      is_read: z.boolean().describe('True to mark as read, false to mark as unread'),
    },
    async ({ id, is_read }) => {
      await markMessageRead(id, is_read);
      return { content: [{ type: 'text', text: `Message marked as ${is_read ? 'read' : 'unread'}.` }] };
    },
  );

  // ── outlook_flag ─────────────────────────────────────────────────────────
  server.tool(
    'outlook_flag',
    'Flag or unflag an email.',
    {
      id: z.string().describe('Message ID'),
      status: z.enum(['Flagged', 'Complete', 'NotFlagged']).describe('Flag status to set'),
    },
    async ({ id, status }) => {
      await flagMessage(id, status);
      return { content: [{ type: 'text', text: `Message flag set to: ${status}.` }] };
    },
  );

  // ── outlook_move_email ───────────────────────────────────────────────────
  server.tool(
    'outlook_move_email',
    'Move an email to a different folder.',
    {
      id: z.string().describe('Message ID to move'),
      destination_folder: z.string().describe('Destination folder name or ID (e.g. Archive, DeletedItems, Inbox, or a folder ID)'),
    },
    async ({ id, destination_folder }) => {
      const moved = await moveMessage(id, destination_folder);
      return { content: [{ type: 'text', text: `Message moved to ${destination_folder}.\nNew ID: ${moved.Id}` }] };
    },
  );

  // ── outlook_delete_email ─────────────────────────────────────────────────
  server.tool(
    'outlook_delete_email',
    'Delete an email (moves to Deleted Items).',
    {
      id: z.string().describe('Message ID to delete'),
    },
    async ({ id }) => {
      await deleteMessage(id);
      return { content: [{ type: 'text', text: 'Message deleted.' }] };
    },
  );

  // ── outlook_batch ────────────────────────────────────────────────────────
  server.tool(
    'outlook_batch',
    'Run one bulk action on many emails at once. Far faster than calling the single item tools repeatedly. Returns a per id success and failure summary.',
    {
      action: z.enum(['mark_read', 'mark_unread', 'delete', 'flag', 'unflag', 'move'])
        .describe('Action to apply to every ID. "move" requires destination_folder.'),
      ids: z.array(z.string()).min(1).max(200).describe('Message IDs to act on (1 to 200)'),
      destination_folder: z.string().optional().describe('Required when action is "move": folder name or ID (e.g. Archive, DeletedItems, or a folder ID)'),
    },
    async ({ action, ids, destination_folder }) => {
      if (action === 'move' && !destination_folder) {
        return { content: [{ type: 'text', text: 'destination_folder is required when action is "move".' }] };
      }

      const opFor = (id: string): Promise<unknown> => {
        switch (action) {
          case 'mark_read': return markMessageRead(id, true);
          case 'mark_unread': return markMessageRead(id, false);
          case 'delete': return deleteMessage(id);
          case 'flag': return flagMessage(id, 'Flagged');
          case 'unflag': return flagMessage(id, 'NotFlagged');
          case 'move': return moveMessage(id, destination_folder!);
        }
      };

      const { succeeded, failed } = await runBatch(ids, opFor);
      const summary = { action, total: ids.length, succeeded, failed: failed.length, failures: failed };
      return { content: [{ type: 'text', text: JSON.stringify(summary, null, 2) }] };
    },
  );

  // ── outlook_search_emails ────────────────────────────────────────────────
  server.tool(
    'outlook_search_emails',
    'Search emails by keyword across subject and body, with optional received date range and pagination. Omit query to list everything in a date range.',
    {
      query: z.string().optional().describe('Search query (keywords, sender name, subject, etc.). Optional: leave empty to list everything that matches the date range.'),
      top: z.number().int().min(1).max(50).optional().describe('Page size (default 20, max 50)'),
      start_date: z.string().optional().describe('Only emails received on or after this date. ISO date (YYYY-MM-DD) or datetime.'),
      end_date: z.string().optional().describe('Only emails received on or before this date. ISO date (YYYY-MM-DD) or datetime.'),
      folder: z.string().optional().describe('Folder to search (default Inbox). E.g. Inbox, Archive, SentItems.'),
      skip_token: z.string().optional().describe('To get the next page, pass the next_skip_token returned by a previous call.'),
    },
    async ({ query, top, start_date, end_date, folder, skip_token }) => {
      const { messages, nextSkipToken } = await searchMessages({
        query,
        top: top ?? 20,
        startDate: start_date,
        endDate: end_date,
        folder,
        skipToken: skip_token,
      });

      let text = formatMessageList(messages);
      if (nextSkipToken) {
        text += `\n\nMore results available. To get the next page, call outlook_search_emails again with the same query and:\nskip_token: ${nextSkipToken}`;
      }
      return { content: [{ type: 'text', text }] };
    },
  );

  // ── outlook_list_folders ─────────────────────────────────────────────────
  server.tool(
    'outlook_list_folders',
    'List all mail folders with unread counts.',
    {},
    async () => {
      const folders = await listFolders();
      const text = folders
        .map(f => `${f.DisplayName} (ID: ${f.Id}) | Unread: ${f.UnreadItemCount} / ${f.TotalItemCount}`)
        .join('\n');
      return { content: [{ type: 'text', text: text || 'No folders found.' }] };
    },
  );

  // ── outlook_update_draft ─────────────────────────────────────────────────
  server.tool(
    'outlook_update_draft',
    'Edit an existing draft: change its subject, recipients, body, or importance. Useful for tweaking a reply or forward draft before sending it with outlook_send_draft. Only the fields you pass are changed. Body follows the same HTML rules as sending.',
    {
      id: z.string().describe('Draft message ID'),
      subject: z.string().optional().describe('New subject'),
      body: z.string().optional().describe('New body. Prefer HTML; plain text newlines auto convert to <br>.'),
      to: z.array(z.string().email()).optional().describe('Replace the To recipients'),
      cc: z.array(z.string().email()).optional().describe('Replace the CC recipients'),
      bcc: z.array(z.string().email()).optional().describe('Replace the BCC recipients'),
      importance: z.enum(['Low', 'Normal', 'High']).optional(),
    },
    async ({ id, subject, body, to, cc, bcc, importance }) => {
      const draft = await updateDraft(id, { subject, body, to, cc, bcc, importance });
      return { content: [{ type: 'text', text: `Draft updated.\nID: ${draft.Id}\nSubject: ${draft.Subject}` }] };
    },
  );

  // ── outlook_add_attachment ───────────────────────────────────────────────
  server.tool(
    'outlook_add_attachment',
    'Attach a local file to an existing draft (including a reply or forward draft). Read the file from disk and add it. Combine with outlook_create_reply_draft or outlook_create_draft, then outlook_send_draft.',
    {
      message_id: z.string().describe('Draft message ID to attach to'),
      file_path: z.string().describe('Local file path to attach'),
    },
    async ({ message_id, file_path }) => {
      const attachment = await fileToAttachment(file_path);
      await addAttachment(message_id, attachment);
      return { content: [{ type: 'text', text: `Attached ${attachment.name} to draft ${message_id}.` }] };
    },
  );

  // ── outlook_list_attachments ─────────────────────────────────────────────
  server.tool(
    'outlook_list_attachments',
    'List the attachments on an email (name, type, size, and attachment ID for downloading).',
    {
      message_id: z.string().describe('Message ID'),
    },
    async ({ message_id }) => {
      const items = await listAttachments(message_id);
      if (items.length === 0) return { content: [{ type: 'text', text: 'No attachments.' }] };
      const text = items.map((a: Attachment) => `- ${a.Name} (${a.ContentType}, ${Math.round(a.Size / 1024)}KB)${a.IsInline ? ' [inline]' : ''}\n  ID: ${a.Id}`).join('\n');
      return { content: [{ type: 'text', text }] };
    },
  );

  // ── outlook_save_attachment ──────────────────────────────────────────────
  server.tool(
    'outlook_save_attachment',
    'Download an email attachment and save it to a local file path. Use outlook_list_attachments first to get the attachment ID.',
    {
      message_id: z.string().describe('Message ID the attachment belongs to'),
      attachment_id: z.string().describe('Attachment ID (from outlook_list_attachments)'),
      output_path: z.string().describe('Local file path to write the attachment to'),
    },
    async ({ message_id, attachment_id, output_path }) => {
      const att = await getAttachmentContent(message_id, attachment_id);
      if (!att.ContentBytes) {
        return { content: [{ type: 'text', text: `Attachment ${att.Name} has no downloadable content (it may be an item or reference attachment).` }] };
      }
      await writeFile(output_path, Buffer.from(att.ContentBytes, 'base64'));
      return { content: [{ type: 'text', text: `Saved ${att.Name} (${Math.round(att.Size / 1024)}KB) to ${output_path}.` }] };
    },
  );

  // ── outlook_get_conversation ─────────────────────────────────────────────
  server.tool(
    'outlook_get_conversation',
    'Get every message in a conversation/thread, oldest first, by conversation ID. The conversation ID comes from any message in the thread (outlook_get_email / outlook_list_emails). Useful for reading a whole back-and-forth before replying.',
    {
      conversation_id: z.string().describe('Conversation ID shared by all messages in the thread'),
      top: z.number().int().min(1).max(100).optional().describe('Max messages to return (default 50)'),
    },
    async ({ conversation_id, top }) => {
      const messages = await getConversation(conversation_id, top ?? 50);
      return { content: [{ type: 'text', text: formatMessageList(messages) }] };
    },
  );

  // ── outlook_set_categories ───────────────────────────────────────────────
  server.tool(
    'outlook_set_categories',
    'Set the colour categories (labels) on an email. This replaces the existing categories with the list you pass; pass an empty list to clear them.',
    {
      message_id: z.string().describe('Message ID'),
      categories: z.array(z.string()).describe('Category names to set (replaces existing). Empty list clears them.'),
    },
    async ({ message_id, categories }) => {
      await setCategories(message_id, categories);
      const text = categories.length ? `Categories set: ${categories.join(', ')}.` : 'Categories cleared.';
      return { content: [{ type: 'text', text }] };
    },
  );

  // ── outlook_create_folder ────────────────────────────────────────────────
  server.tool(
    'outlook_create_folder',
    'Create a new mail folder, optionally nested under a parent folder.',
    {
      name: z.string().describe('Folder display name'),
      parent_folder_id: z.string().optional().describe('Parent folder ID to nest under (defaults to top level)'),
    },
    async ({ name, parent_folder_id }) => {
      const folder = await createFolder(name, parent_folder_id);
      return { content: [{ type: 'text', text: `Folder created.\nID: ${folder.Id}\nName: ${folder.DisplayName}` }] };
    },
  );

  // ── outlook_rename_folder ────────────────────────────────────────────────
  server.tool(
    'outlook_rename_folder',
    'Rename an existing mail folder.',
    {
      id: z.string().describe('Folder ID to rename'),
      name: z.string().describe('New folder display name'),
    },
    async ({ id, name }) => {
      const folder = await renameFolder(id, name);
      return { content: [{ type: 'text', text: `Folder renamed to ${folder.DisplayName}.` }] };
    },
  );

  // ── outlook_delete_folder ────────────────────────────────────────────────
  server.tool(
    'outlook_delete_folder',
    'Delete a mail folder and everything in it. This is destructive: confirm with the user and double check the folder ID before calling.',
    {
      id: z.string().describe('Folder ID to delete'),
    },
    async ({ id }) => {
      await deleteFolder(id);
      return { content: [{ type: 'text', text: 'Folder deleted.' }] };
    },
  );
}
