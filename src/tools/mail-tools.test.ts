import { describe, it, expect, vi, beforeEach } from 'vitest';
import * as mail from '../api/mail.js';
import { writeFile } from 'node:fs/promises';
import { registerMailTools } from './mail-tools.js';

vi.mock('node:fs/promises', () => ({ writeFile: vi.fn() }));
vi.mock('../api/mail.js', () => ({
  listMessages: vi.fn(), getMessage: vi.fn(), sendEmail: vi.fn(), createDraft: vi.fn(),
  replyToMessage: vi.fn(), createReplyDraft: vi.fn(), createForwardDraft: vi.fn(),
  forwardMessage: vi.fn(), markMessageRead: vi.fn(), flagMessage: vi.fn(), moveMessage: vi.fn(),
  deleteMessage: vi.fn(), searchMessages: vi.fn(), listFolders: vi.fn(), getUnreadMessages: vi.fn(),
  sendDraft: vi.fn(), updateDraft: vi.fn(), fileToAttachment: vi.fn(), listAttachments: vi.fn(),
  getAttachmentContent: vi.fn(), addAttachment: vi.fn(), getConversation: vi.fn(),
  setCategories: vi.fn(), createFolder: vi.fn(), renameFolder: vi.fn(), deleteFolder: vi.fn(),
}));

type Reg = { schema: any; handler: (a: any) => Promise<any> };
let tools: Map<string, Reg>;
function setup() {
  tools = new Map();
  const server = { tool: (n: string, _d: string, s: any, h: any) => tools.set(n, { schema: s, handler: h }) } as any;
  registerMailTools(server);
}
const text = (r: any) => r.content[0].text;

const richMsg = {
  Id: 'm1', Subject: 'Hello', From: { EmailAddress: { Name: 'Jane', Address: 'jane@x.com' } },
  ToRecipients: [{ EmailAddress: { Address: 'me@x.com' } }], ReceivedDateTime: '2024-01-01',
  IsRead: true, HasAttachments: true, Flag: { FlagStatus: 'Flagged' }, WebLink: 'http://w',
  Body: { ContentType: 'HTML', Content: 'x'.repeat(6000) },
  Attachments: [{ Name: 'a.pdf', ContentType: 'application/pdf', Size: 2048 }],
};
const leanMsg = {
  Id: 'm2', Subject: 'Hi', From: undefined, IsRead: false, HasAttachments: false,
  Flag: { FlagStatus: 'NotFlagged' }, SentDateTime: '2024-02-02', BodyPreview: 'preview text',
};

beforeEach(() => { vi.clearAllMocks(); setup(); });

describe('outlook_list_emails', () => {
  it('lists with defaults and formats lean message', async () => {
    vi.mocked(mail.listMessages).mockResolvedValue([leanMsg] as any);
    const r = await tools.get('outlook_list_emails')!.handler({});
    expect(mail.listMessages).toHaveBeenCalledWith({ folder: 'Inbox', top: 20, filter: undefined });
    expect(text(r)).toContain('Unknown');
    expect(text(r)).toContain('Preview: preview text');
  });
  it('unread_only + mark_read marks each', async () => {
    vi.mocked(mail.listMessages).mockResolvedValue([richMsg, leanMsg] as any);
    vi.mocked(mail.markMessageRead).mockResolvedValue(undefined as any);
    const r = await tools.get('outlook_list_emails')!.handler({ folder: 'Archive', top: 5, unread_only: true, mark_read: true });
    expect(mail.listMessages).toHaveBeenCalledWith({ folder: 'Archive', top: 5, filter: 'IsRead eq false' });
    expect(mail.markMessageRead).toHaveBeenCalledTimes(2);
    expect(text(r)).toContain('Flag: Flagged');
  });
  it('empty list', async () => {
    vi.mocked(mail.listMessages).mockResolvedValue([] as any);
    expect(text(await tools.get('outlook_list_emails')!.handler({}))).toBe('No messages found.');
  });
});

describe('outlook_get_email', () => {
  it('reads full + attachments, marks read when unread', async () => {
    vi.mocked(mail.getMessage).mockResolvedValue({ ...richMsg, IsRead: false } as any);
    vi.mocked(mail.markMessageRead).mockResolvedValue(undefined as any);
    const r = await tools.get('outlook_get_email')!.handler({ id: 'm1', include_attachments: true, mark_read: true });
    expect(mail.getMessage).toHaveBeenCalledWith('m1', true);
    expect(mail.markMessageRead).toHaveBeenCalledWith('m1');
    expect(text(r)).toContain('Attachments:');
    expect(text(r)).toContain('a.pdf');
  });
  it('no mark when already read, no attachments', async () => {
    vi.mocked(mail.getMessage).mockResolvedValue({ ...richMsg, Attachments: [] } as any);
    const r = await tools.get('outlook_get_email')!.handler({ id: 'm1', mark_read: true });
    expect(mail.getMessage).toHaveBeenCalledWith('m1', false);
    expect(mail.markMessageRead).not.toHaveBeenCalled();
    expect(text(r)).not.toContain('- a.pdf');
  });
});

describe('outlook_get_unread', () => {
  it('plural', async () => {
    vi.mocked(mail.getUnreadMessages).mockResolvedValue([richMsg, leanMsg] as any);
    expect(text(await tools.get('outlook_get_unread')!.handler({ top: 5 }))).toContain('2 unread emails');
  });
  it('singular default top', async () => {
    vi.mocked(mail.getUnreadMessages).mockResolvedValue([richMsg] as any);
    const r = await tools.get('outlook_get_unread')!.handler({});
    expect(mail.getUnreadMessages).toHaveBeenCalledWith(10);
    expect(text(r)).toContain('1 unread email:');
  });
});

describe('outlook_send_email', () => {
  it('sends with attachments', async () => {
    vi.mocked(mail.fileToAttachment).mockResolvedValue({ name: 'f.txt' } as any);
    vi.mocked(mail.sendEmail).mockResolvedValue(undefined as any);
    const r = await tools.get('outlook_send_email')!.handler({ to: ['a@x.com'], subject: 's', body: 'b', attachments: ['/p/f.txt'] });
    expect(vi.mocked(mail.fileToAttachment).mock.calls[0][0]).toBe('/p/f.txt');
    expect(text(r)).toContain('with 1 attachment(s)');
  });
  it('sends without attachments', async () => {
    vi.mocked(mail.sendEmail).mockResolvedValue(undefined as any);
    const r = await tools.get('outlook_send_email')!.handler({ to: ['a@x.com', 'b@x.com'], subject: 's', body: 'b' });
    expect(text(r)).toBe('Email sent to a@x.com, b@x.com.');
  });
});

describe('outlook_create_draft', () => {
  it('creates with attachments', async () => {
    vi.mocked(mail.fileToAttachment).mockResolvedValue({ name: 'f' } as any);
    vi.mocked(mail.createDraft).mockResolvedValue({ Id: 'd1', Subject: 's' } as any);
    const r = await tools.get('outlook_create_draft')!.handler({ to: ['a@x.com'], subject: 's', body: 'b', attachments: ['/f'] });
    expect(text(r)).toContain('Draft created.');
    expect(text(r)).toContain('d1');
  });
  it('creates without attachments', async () => {
    vi.mocked(mail.createDraft).mockResolvedValue({ Id: 'd2', Subject: 's' } as any);
    const r = await tools.get('outlook_create_draft')!.handler({ to: ['a@x.com'], subject: 's', body: 'b' });
    expect(text(r)).toContain('d2');
  });
});

describe('simple action tools', () => {
  it('send_draft', async () => {
    vi.mocked(mail.sendDraft).mockResolvedValue(undefined as any);
    expect(text(await tools.get('outlook_send_draft')!.handler({ id: 'd' }))).toContain('Draft sent');
    expect(mail.sendDraft).toHaveBeenCalledWith('d');
  });
  it('reply default and reply_all', async () => {
    vi.mocked(mail.replyToMessage).mockResolvedValue(undefined as any);
    await tools.get('outlook_reply')!.handler({ id: 'm', body: 'b' });
    expect(mail.replyToMessage).toHaveBeenCalledWith('m', 'b', false);
    await tools.get('outlook_reply')!.handler({ id: 'm', body: 'b', reply_all: true });
    expect(mail.replyToMessage).toHaveBeenCalledWith('m', 'b', true);
  });
  it('create_reply_draft', async () => {
    vi.mocked(mail.createReplyDraft).mockResolvedValue({ Id: 'r', Subject: 's' } as any);
    expect(text(await tools.get('outlook_create_reply_draft')!.handler({ id: 'm', body: 'b', reply_all: true }))).toContain('Reply draft created');
  });
  it('create_forward_draft with and without comment', async () => {
    vi.mocked(mail.createForwardDraft).mockResolvedValue({ Id: 'f', Subject: 's' } as any);
    await tools.get('outlook_create_forward_draft')!.handler({ id: 'm', to: ['a@x.com'], comment: 'hi' });
    expect(mail.createForwardDraft).toHaveBeenCalledWith('m', 'hi', ['a@x.com']);
    await tools.get('outlook_create_forward_draft')!.handler({ id: 'm' });
    expect(mail.createForwardDraft).toHaveBeenCalledWith('m', '', undefined);
  });
  it('forward', async () => {
    vi.mocked(mail.forwardMessage).mockResolvedValue(undefined as any);
    expect(text(await tools.get('outlook_forward')!.handler({ id: 'm', to: ['a@x.com'], comment: 'c' }))).toContain('Forwarded to a@x.com');
  });
  it('mark_read true/false', async () => {
    vi.mocked(mail.markMessageRead).mockResolvedValue(undefined as any);
    expect(text(await tools.get('outlook_mark_read')!.handler({ id: 'm', is_read: true }))).toContain('marked as read');
    expect(text(await tools.get('outlook_mark_read')!.handler({ id: 'm', is_read: false }))).toContain('marked as unread');
  });
  it('flag', async () => {
    vi.mocked(mail.flagMessage).mockResolvedValue(undefined as any);
    expect(text(await tools.get('outlook_flag')!.handler({ id: 'm', status: 'Complete' }))).toContain('Complete');
  });
  it('move', async () => {
    vi.mocked(mail.moveMessage).mockResolvedValue({ Id: 'new' } as any);
    expect(text(await tools.get('outlook_move_email')!.handler({ id: 'm', destination_folder: 'Archive' }))).toContain('New ID: new');
  });
  it('delete', async () => {
    vi.mocked(mail.deleteMessage).mockResolvedValue(undefined as any);
    expect(text(await tools.get('outlook_delete_email')!.handler({ id: 'm' }))).toBe('Message deleted.');
  });
});

describe('outlook_batch', () => {
  it('move without destination errors', async () => {
    expect(text(await tools.get('outlook_batch')!.handler({ action: 'move', ids: ['a'] }))).toContain('destination_folder is required');
  });
  it('mark_read success', async () => {
    vi.mocked(mail.markMessageRead).mockResolvedValue(undefined as any);
    const s = JSON.parse(text(await tools.get('outlook_batch')!.handler({ action: 'mark_read', ids: ['a', 'b'] })));
    expect(s.succeeded).toBe(2);
    expect(mail.markMessageRead).toHaveBeenCalledWith('a', true);
  });
  it('mark_unread', async () => {
    vi.mocked(mail.markMessageRead).mockResolvedValue(undefined as any);
    await tools.get('outlook_batch')!.handler({ action: 'mark_unread', ids: ['a'] });
    expect(mail.markMessageRead).toHaveBeenCalledWith('a', false);
  });
  it('delete with Error failure', async () => {
    vi.mocked(mail.deleteMessage).mockRejectedValueOnce(new Error('boom')).mockResolvedValueOnce(undefined as any);
    const s = JSON.parse(text(await tools.get('outlook_batch')!.handler({ action: 'delete', ids: ['a', 'b'] })));
    expect(s.succeeded).toBe(1);
    expect(s.failures[0]).toEqual({ id: 'a', error: 'boom' });
  });
  it('flag with non-Error rejection', async () => {
    vi.mocked(mail.flagMessage).mockRejectedValue('strfail' as any);
    const s = JSON.parse(text(await tools.get('outlook_batch')!.handler({ action: 'flag', ids: ['a'] })));
    expect(s.failures[0].error).toBe('strfail');
  });
  it('unflag', async () => {
    vi.mocked(mail.flagMessage).mockResolvedValue(undefined as any);
    await tools.get('outlook_batch')!.handler({ action: 'unflag', ids: ['a'] });
    expect(mail.flagMessage).toHaveBeenCalledWith('a', 'NotFlagged');
  });
  it('move with destination', async () => {
    vi.mocked(mail.moveMessage).mockResolvedValue({ Id: 'n' } as any);
    const s = JSON.parse(text(await tools.get('outlook_batch')!.handler({ action: 'move', ids: ['a'], destination_folder: 'Archive' })));
    expect(s.succeeded).toBe(1);
    expect(mail.moveMessage).toHaveBeenCalledWith('a', 'Archive');
  });
});

describe('outlook_search_emails', () => {
  it('with next skip token', async () => {
    vi.mocked(mail.searchMessages).mockResolvedValue({ messages: [richMsg], nextSkipToken: 'tok' } as any);
    const r = await tools.get('outlook_search_emails')!.handler({ query: 'q' });
    expect(mail.searchMessages).toHaveBeenCalledWith({ query: 'q', top: 20, startDate: undefined, endDate: undefined, folder: undefined, skipToken: undefined });
    expect(text(r)).toContain('skip_token: tok');
  });
  it('without next token, custom args', async () => {
    vi.mocked(mail.searchMessages).mockResolvedValue({ messages: [], nextSkipToken: undefined } as any);
    const r = await tools.get('outlook_search_emails')!.handler({ query: 'q', top: 5, start_date: 's', end_date: 'e', folder: 'Inbox', skip_token: 'k' });
    expect(text(r)).toBe('No messages found.');
  });
});

describe('folders', () => {
  it('list folders', async () => {
    vi.mocked(mail.listFolders).mockResolvedValue([{ DisplayName: 'Inbox', Id: 'i', UnreadItemCount: 2, TotalItemCount: 10 }] as any);
    expect(text(await tools.get('outlook_list_folders')!.handler({}))).toContain('Inbox (ID: i) | Unread: 2 / 10');
  });
  it('list folders empty', async () => {
    vi.mocked(mail.listFolders).mockResolvedValue([] as any);
    expect(text(await tools.get('outlook_list_folders')!.handler({}))).toBe('No folders found.');
  });
  it('create folder', async () => {
    vi.mocked(mail.createFolder).mockResolvedValue({ Id: 'f', DisplayName: 'New' } as any);
    await tools.get('outlook_create_folder')!.handler({ name: 'New', parent_folder_id: 'p' });
    expect(mail.createFolder).toHaveBeenCalledWith('New', 'p');
  });
  it('rename folder', async () => {
    vi.mocked(mail.renameFolder).mockResolvedValue({ DisplayName: 'X' } as any);
    expect(text(await tools.get('outlook_rename_folder')!.handler({ id: 'f', name: 'X' }))).toContain('renamed to X');
  });
  it('delete folder', async () => {
    vi.mocked(mail.deleteFolder).mockResolvedValue(undefined as any);
    expect(text(await tools.get('outlook_delete_folder')!.handler({ id: 'f' }))).toBe('Folder deleted.');
  });
});

describe('drafts/attachments/misc', () => {
  it('update_draft', async () => {
    vi.mocked(mail.updateDraft).mockResolvedValue({ Id: 'd', Subject: 's' } as any);
    const r = await tools.get('outlook_update_draft')!.handler({ id: 'd', subject: 's', body: 'b', to: ['a@x.com'], cc: [], bcc: [], importance: 'High' });
    expect(text(r)).toContain('Draft updated.');
  });
  it('add_attachment', async () => {
    vi.mocked(mail.fileToAttachment).mockResolvedValue({ name: 'a.txt' } as any);
    vi.mocked(mail.addAttachment).mockResolvedValue(undefined as any);
    expect(text(await tools.get('outlook_add_attachment')!.handler({ message_id: 'm', file_path: '/a.txt' }))).toContain('Attached a.txt');
  });
  it('list_attachments with items', async () => {
    vi.mocked(mail.listAttachments).mockResolvedValue([
      { Name: 'a', ContentType: 't', Size: 2048, Id: 'x', IsInline: true },
      { Name: 'b', ContentType: 't', Size: 1024, Id: 'y' },
    ] as any);
    const r = text(await tools.get('outlook_list_attachments')!.handler({ message_id: 'm' }));
    expect(r).toContain('[inline]');
    expect(r).toContain('ID: y');
  });
  it('list_attachments empty', async () => {
    vi.mocked(mail.listAttachments).mockResolvedValue([] as any);
    expect(text(await tools.get('outlook_list_attachments')!.handler({ message_id: 'm' }))).toBe('No attachments.');
  });
  it('save_attachment writes file', async () => {
    vi.mocked(mail.getAttachmentContent).mockResolvedValue({ Name: 'a', Size: 1024, ContentBytes: Buffer.from('hi').toString('base64') } as any);
    const r = await tools.get('outlook_save_attachment')!.handler({ message_id: 'm', attachment_id: 'a', output_path: '/out' });
    expect(writeFile).toHaveBeenCalled();
    expect(text(r)).toContain('Saved a');
  });
  it('save_attachment no content', async () => {
    vi.mocked(mail.getAttachmentContent).mockResolvedValue({ Name: 'a', Size: 0, ContentBytes: null } as any);
    const r = await tools.get('outlook_save_attachment')!.handler({ message_id: 'm', attachment_id: 'a', output_path: '/out' });
    expect(text(r)).toContain('no downloadable content');
    expect(writeFile).not.toHaveBeenCalled();
  });
  it('get_conversation', async () => {
    vi.mocked(mail.getConversation).mockResolvedValue([richMsg] as any);
    await tools.get('outlook_get_conversation')!.handler({ conversation_id: 'c' });
    expect(mail.getConversation).toHaveBeenCalledWith('c', 50);
  });
  it('set_categories set and clear', async () => {
    vi.mocked(mail.setCategories).mockResolvedValue(undefined as any);
    expect(text(await tools.get('outlook_set_categories')!.handler({ message_id: 'm', categories: ['A', 'B'] }))).toContain('Categories set: A, B');
    expect(text(await tools.get('outlook_set_categories')!.handler({ message_id: 'm', categories: [] }))).toBe('Categories cleared.');
  });
});
