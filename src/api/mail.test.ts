import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('./client.js', () => ({
  owaGet: vi.fn(),
  owaPost: vi.fn(),
  owaPatch: vi.fn(),
  owaDelete: vi.fn(),
}));

vi.mock('node:fs/promises', () => ({
  readFile: vi.fn(),
}));

import { owaGet, owaPost, owaPatch, owaDelete } from './client.js';
import { readFile } from 'node:fs/promises';
import {
  looksLikeHtml, toHtmlBody, fileToAttachment, listMessages, getMessage,
  sendEmail, createDraft, replyToMessage, createReplyDraft, createForwardDraft,
  forwardMessage, markMessageRead, flagMessage, moveMessage, deleteMessage,
  searchMessages, listFolders, getFolder, getUnreadMessages, sendDraft,
  updateDraft, listAttachments, getAttachmentContent, addAttachment,
  getConversation, setCategories, createFolder, renameFolder, deleteFolder,
} from './mail.js';

const mGet = vi.mocked(owaGet);
const mPost = vi.mocked(owaPost);
const mPatch = vi.mocked(owaPatch);
const mDelete = vi.mocked(owaDelete);
const mReadFile = vi.mocked(readFile);

beforeEach(() => vi.clearAllMocks());

describe('looksLikeHtml', () => {
  it('detects real html tags', () => {
    expect(looksLikeHtml('<br>')).toBe(true);
    expect(looksLikeHtml('<div>hi</div>')).toBe(true);
  });
  it('treats plain text and lone angle brackets as not html', () => {
    expect(looksLikeHtml('Integrations > Vincere')).toBe(false);
    expect(looksLikeHtml('just text')).toBe(false);
  });
});

describe('toHtmlBody', () => {
  it('passes through html', () => {
    const html = '<div>Hi</div><br>';
    expect(toHtmlBody(html)).toBe(html);
  });
  it('converts newlines and escapes plain text', () => {
    expect(toHtmlBody('one\ntwo')).toBe('one<br>two');
    expect(toHtmlBody('one\r\ntwo')).toBe('one<br>two');
    expect(toHtmlBody('one\rtwo')).toBe('one<br>two');
    expect(toHtmlBody('a < b & c > d')).toBe('a &lt; b &amp; c &gt; d');
  });
});

describe('fileToAttachment', () => {
  it('reads a known extension and base64 encodes', async () => {
    mReadFile.mockResolvedValue(Buffer.from('hello'));
    const a = await fileToAttachment('/tmp/report.pdf');
    expect(mReadFile).toHaveBeenCalledWith('/tmp/report.pdf');
    expect(a).toEqual({
      name: 'report.pdf',
      contentType: 'application/pdf',
      contentBytes: Buffer.from('hello').toString('base64'),
    });
  });
  it('falls back to octet-stream for unknown extension', async () => {
    mReadFile.mockResolvedValue(Buffer.from('x'));
    const a = await fileToAttachment('/tmp/file.unknownext');
    expect(a.contentType).toBe('application/octet-stream');
    expect(a.name).toBe('file.unknownext');
  });
});

describe('listMessages', () => {
  it('uses defaults', async () => {
    mGet.mockResolvedValue({ value: [{ Id: '1' }] });
    const res = await listMessages();
    expect(res).toEqual([{ Id: '1' }]);
    const [path, params] = mGet.mock.calls[0] as [string, Record<string, string>];
    expect(path).toBe('/MailFolders/Inbox/messages');
    expect(params['$top']).toBe('20');
    expect(params['$orderby']).toBe('ReceivedDateTime desc');
    expect(params['$select']).toContain('Id');
    expect(params['$skip']).toBeUndefined();
  });

  it('applies folder, top, skip, select, orderBy and filter', async () => {
    mGet.mockResolvedValue({ value: [] });
    await listMessages({ folder: 'Sent', top: 5, skip: 10, select: ['Id'], orderBy: 'Subject asc', filter: "IsRead eq false" });
    const [path, params] = mGet.mock.calls[0] as [string, Record<string, string>];
    expect(path).toBe('/MailFolders/Sent/messages');
    expect(params['$top']).toBe('5');
    expect(params['$skip']).toBe('10');
    expect(params['$select']).toBe('Id');
    expect(params['$orderby']).toBe('Subject asc');
    expect(params['$filter']).toBe('IsRead eq false');
  });

  it('uses $search and drops ordering/filter when searching', async () => {
    mGet.mockResolvedValue({ value: [] });
    await listMessages({ search: 'invoice', filter: 'x', orderBy: 'y' });
    const params = mGet.mock.calls[0][1] as Record<string, string>;
    expect(params['$search']).toBe('"invoice"');
    expect(params['$orderby']).toBeUndefined();
    expect(params['$filter']).toBeUndefined();
  });
});

describe('getMessage', () => {
  it('without attachments', async () => {
    mGet.mockResolvedValue({ Id: '1' });
    await getMessage('1');
    expect(mGet).toHaveBeenCalledWith('/messages/1', undefined);
  });
  it('with attachments expand', async () => {
    mGet.mockResolvedValue({ Id: '1' });
    await getMessage('1', true);
    expect(mGet).toHaveBeenCalledWith('/messages/1', { '$expand': 'Attachments' });
  });
});

describe('sendEmail', () => {
  it('sends with defaults (HTML body, save to sent)', async () => {
    mPost.mockResolvedValue(undefined);
    await sendEmail({ to: ['a@b.com'], subject: 'Hi', body: 'line1\nline2' });
    const [path, body] = mPost.mock.calls[0] as [string, any];
    expect(path).toBe('/sendmail');
    expect(body.Message.Body).toEqual({ ContentType: 'HTML', Content: 'line1<br>line2' });
    expect(body.Message.ToRecipients).toEqual([{ EmailAddress: { Address: 'a@b.com' } }]);
    expect(body.Message.Importance).toBe('Normal');
    expect(body.Message.CcRecipients).toBeUndefined();
    expect(body.Message.BccRecipients).toBeUndefined();
    expect(body.Message.Attachments).toBeUndefined();
    expect(body.SaveToSentItems).toBe(true);
  });

  it('sends with cc, bcc, attachments, importance, text body, no save', async () => {
    mPost.mockResolvedValue(undefined);
    await sendEmail({
      to: ['a@b.com'], cc: ['c@b.com'], bcc: ['d@b.com'],
      subject: 'S', body: 'plain\ntext', bodyType: 'Text', importance: 'High',
      attachments: [{ name: 'f.txt', contentType: 'text/plain', contentBytes: 'AAA' }],
      saveToSentItems: false,
    });
    const body = mPost.mock.calls[0][1] as any;
    expect(body.Message.Body).toEqual({ ContentType: 'Text', Content: 'plain\ntext' });
    expect(body.Message.CcRecipients).toEqual([{ EmailAddress: { Address: 'c@b.com' } }]);
    expect(body.Message.BccRecipients).toEqual([{ EmailAddress: { Address: 'd@b.com' } }]);
    expect(body.Message.Importance).toBe('High');
    expect(body.Message.Attachments[0]['@odata.type']).toBe('#Microsoft.OutlookServices.FileAttachment');
    expect(body.Message.Attachments[0].Name).toBe('f.txt');
    expect(body.SaveToSentItems).toBe(false);
  });
});

describe('createDraft', () => {
  it('creates draft with defaults', async () => {
    mPost.mockResolvedValue({ Id: 'd1' });
    const res = await createDraft({ to: ['a@b.com'], subject: 'S', body: '<b>hi</b>' });
    expect(res).toEqual({ Id: 'd1' });
    const [path, body] = mPost.mock.calls[0] as [string, any];
    expect(path).toBe('/messages');
    expect(body.Body).toEqual({ ContentType: 'HTML', Content: '<b>hi</b>' });
    expect(body.CcRecipients).toBeUndefined();
  });
  it('creates draft with cc/bcc/text/attachments', async () => {
    mPost.mockResolvedValue({ Id: 'd2' });
    await createDraft({
      to: ['a@b.com'], cc: ['c@b.com'], bcc: ['d@b.com'], subject: 'S', body: 'x',
      bodyType: 'Text', attachments: [{ name: 'n', contentType: 't', contentBytes: 'b' }],
    });
    const body = mPost.mock.calls[0][1] as any;
    expect(body.Body.ContentType).toBe('Text');
    expect(body.CcRecipients).toBeDefined();
    expect(body.BccRecipients).toBeDefined();
    expect(body.Attachments).toHaveLength(1);
  });
});

describe('replyToMessage', () => {
  it('reply', async () => {
    mPost.mockResolvedValue(undefined);
    await replyToMessage('1', 'thanks');
    expect(mPost).toHaveBeenCalledWith('/messages/1/reply', { Comment: 'thanks' });
  });
  it('reply all', async () => {
    mPost.mockResolvedValue(undefined);
    await replyToMessage('1', 'thanks', true);
    expect(mPost).toHaveBeenCalledWith('/messages/1/replyall', { Comment: 'thanks' });
  });
});

describe('createReplyDraft', () => {
  it('reply draft', async () => {
    mPost.mockResolvedValue({ Id: 'r1' });
    await createReplyDraft('1', 'body');
    expect(mPost).toHaveBeenCalledWith('/messages/1/createreply', { Comment: 'body' });
  });
  it('reply all draft', async () => {
    mPost.mockResolvedValue({ Id: 'r2' });
    await createReplyDraft('1', 'body', true);
    expect(mPost).toHaveBeenCalledWith('/messages/1/createreplyall', { Comment: 'body' });
  });
});

describe('createForwardDraft', () => {
  it('default empty body and no recipients', async () => {
    mPost.mockResolvedValue({ Id: 'f1' });
    await createForwardDraft('1');
    expect(mPost).toHaveBeenCalledWith('/messages/1/createforward', { Comment: '' });
  });
  it('with body and recipients', async () => {
    mPost.mockResolvedValue({ Id: 'f2' });
    await createForwardDraft('1', 'see below', ['a@b.com']);
    const body = mPost.mock.calls[0][1] as any;
    expect(body.Comment).toBe('see below');
    expect(body.ToRecipients).toEqual([{ EmailAddress: { Address: 'a@b.com' } }]);
  });
});

describe('forwardMessage', () => {
  it('forwards with comment', async () => {
    mPost.mockResolvedValue(undefined);
    await forwardMessage('1', ['a@b.com'], 'fyi');
    const body = mPost.mock.calls[0][1] as any;
    expect(body.Comment).toBe('fyi');
    expect(body.ToRecipients).toEqual([{ EmailAddress: { Address: 'a@b.com' } }]);
  });
  it('forwards without comment', async () => {
    mPost.mockResolvedValue(undefined);
    await forwardMessage('1', ['a@b.com']);
    expect((mPost.mock.calls[0][1] as any).Comment).toBe('');
  });
});

describe('markMessageRead / flagMessage', () => {
  it('marks read by default', async () => {
    mPatch.mockResolvedValue(undefined);
    await markMessageRead('1');
    expect(mPatch).toHaveBeenCalledWith('/messages/1', { IsRead: true });
  });
  it('marks unread', async () => {
    mPatch.mockResolvedValue(undefined);
    await markMessageRead('1', false);
    expect(mPatch).toHaveBeenCalledWith('/messages/1', { IsRead: false });
  });
  it('flags default', async () => {
    mPatch.mockResolvedValue(undefined);
    await flagMessage('1');
    expect(mPatch).toHaveBeenCalledWith('/messages/1', { Flag: { FlagStatus: 'Flagged' } });
  });
  it('flags complete', async () => {
    mPatch.mockResolvedValue(undefined);
    await flagMessage('1', 'Complete');
    expect(mPatch).toHaveBeenCalledWith('/messages/1', { Flag: { FlagStatus: 'Complete' } });
  });
});

describe('moveMessage / deleteMessage', () => {
  it('moves', async () => {
    mPost.mockResolvedValue({ Id: '2' });
    await moveMessage('1', 'folderX');
    expect(mPost).toHaveBeenCalledWith('/messages/1/move', { DestinationId: 'folderX' });
  });
  it('deletes', async () => {
    mDelete.mockResolvedValue(undefined);
    await deleteMessage('1');
    expect(mDelete).toHaveBeenCalledWith('/messages/1');
  });
});

describe('searchMessages', () => {
  it('throws when no query and no dates', async () => {
    await expect(searchMessages({})).rejects.toThrow('searchMessages needs a query');
  });

  it('searches with query only and no next page', async () => {
    mGet.mockResolvedValue({ value: [{ Id: '1' }] });
    const res = await searchMessages({ query: '  hello  ' });
    expect(res.messages).toEqual([{ Id: '1' }]);
    expect(res.nextSkipToken).toBeUndefined();
    const params = mGet.mock.calls[0][1] as Record<string, string>;
    expect(params['$search']).toBe('"hello"');
  });

  it('combines query and date range into KQL', async () => {
    mGet.mockResolvedValue({ value: [] });
    await searchMessages({ query: 'inv', startDate: '2024-01-01T00:00:00Z', endDate: '2024-02-01', folder: 'Sent', top: 3, skipToken: 'tok' });
    const params = mGet.mock.calls[0][1] as Record<string, string>;
    expect(params['$search']).toBe('"inv AND received>=2024-01-01 AND received<=2024-02-01"');
    expect(params['$top']).toBe('3');
    expect(params['$skiptoken']).toBe('tok');
    expect(mGet.mock.calls[0][0]).toBe('/MailFolders/Sent/messages');
  });

  it('extracts skiptoken from nextLink', async () => {
    mGet.mockResolvedValue({ value: [], '@odata.nextLink': 'https://x/y?$skiptoken=NEXT' });
    const res = await searchMessages({ query: 'a' });
    expect(res.nextSkipToken).toBe('NEXT');
  });

  it('returns undefined skiptoken when nextLink has none', async () => {
    mGet.mockResolvedValue({ value: [], '@odata.nextLink': 'https://x/y?foo=1' });
    const res = await searchMessages({ query: 'a' });
    expect(res.nextSkipToken).toBeUndefined();
  });

  it('returns undefined skiptoken when nextLink is malformed', async () => {
    mGet.mockResolvedValue({ value: [], '@odata.nextLink': 'not a url' });
    const res = await searchMessages({ query: 'a' });
    expect(res.nextSkipToken).toBeUndefined();
  });

  it('works with date range only', async () => {
    mGet.mockResolvedValue({ value: [] });
    await searchMessages({ endDate: '2024-03-03' });
    expect((mGet.mock.calls[0][1] as Record<string, string>)['$search']).toBe('"received<=2024-03-03"');
  });
});

describe('folders', () => {
  it('lists folders', async () => {
    mGet.mockResolvedValue({ value: [{ Id: 'f' }] });
    const res = await listFolders();
    expect(res).toEqual([{ Id: 'f' }]);
    expect(mGet.mock.calls[0][0]).toBe('/mailfolders');
  });
  it('gets a folder', async () => {
    mGet.mockResolvedValue({ Id: 'f' });
    await getFolder('Inbox');
    expect(mGet).toHaveBeenCalledWith('/mailfolders/Inbox');
  });
  it('creates folder at root', async () => {
    mPost.mockResolvedValue({ Id: 'f' });
    await createFolder('Projects');
    expect(mPost).toHaveBeenCalledWith('/mailfolders', { DisplayName: 'Projects' });
  });
  it('creates child folder', async () => {
    mPost.mockResolvedValue({ Id: 'f' });
    await createFolder('Sub', 'parent1');
    expect(mPost).toHaveBeenCalledWith('/mailfolders/parent1/childfolders', { DisplayName: 'Sub' });
  });
  it('renames folder', async () => {
    mPatch.mockResolvedValue({ Id: 'f' });
    await renameFolder('f1', 'New');
    expect(mPatch).toHaveBeenCalledWith('/mailfolders/f1', { DisplayName: 'New' });
  });
  it('deletes folder', async () => {
    mDelete.mockResolvedValue(undefined);
    await deleteFolder('f1');
    expect(mDelete).toHaveBeenCalledWith('/mailfolders/f1');
  });
});

describe('getUnreadMessages', () => {
  it('lists unread with default top', async () => {
    mGet.mockResolvedValue({ value: [] });
    await getUnreadMessages();
    const params = mGet.mock.calls[0][1] as Record<string, string>;
    expect(params['$filter']).toBe('IsRead eq false');
    expect(params['$top']).toBe('10');
  });
  it('honours custom top', async () => {
    mGet.mockResolvedValue({ value: [] });
    await getUnreadMessages(3);
    expect((mGet.mock.calls[0][1] as Record<string, string>)['$top']).toBe('3');
  });
});

describe('sendDraft', () => {
  it('sends a draft', async () => {
    mPost.mockResolvedValue(undefined);
    await sendDraft('d1');
    expect(mPost).toHaveBeenCalledWith('/messages/d1/send', {});
  });
});

describe('updateDraft', () => {
  it('updates all fields with HTML body', async () => {
    mPatch.mockResolvedValue({ Id: 'd1' });
    await updateDraft('d1', {
      subject: 'S', body: 'a\nb', to: ['a@b.com'], cc: ['c@b.com'], bcc: ['d@b.com'], importance: 'Low',
    });
    const patch = mPatch.mock.calls[0][1] as any;
    expect(patch.Subject).toBe('S');
    expect(patch.Body).toEqual({ ContentType: 'HTML', Content: 'a<br>b' });
    expect(patch.ToRecipients).toEqual([{ EmailAddress: { Address: 'a@b.com' } }]);
    expect(patch.CcRecipients).toBeDefined();
    expect(patch.BccRecipients).toBeDefined();
    expect(patch.Importance).toBe('Low');
  });
  it('updates text body', async () => {
    mPatch.mockResolvedValue({ Id: 'd1' });
    await updateDraft('d1', { body: 'a\nb', bodyType: 'Text' });
    expect((mPatch.mock.calls[0][1] as any).Body).toEqual({ ContentType: 'Text', Content: 'a\nb' });
  });
  it('sends empty patch when nothing provided', async () => {
    mPatch.mockResolvedValue({ Id: 'd1' });
    await updateDraft('d1', {});
    expect(mPatch).toHaveBeenCalledWith('/messages/d1', {});
  });
});

describe('attachments', () => {
  it('lists attachments', async () => {
    mGet.mockResolvedValue({ value: [{ Id: 'a' }] });
    const res = await listAttachments('m1');
    expect(res).toEqual([{ Id: 'a' }]);
    expect(mGet.mock.calls[0][0]).toBe('/messages/m1/attachments');
  });
  it('gets attachment content', async () => {
    mGet.mockResolvedValue({ Id: 'a', ContentBytes: 'AAA' });
    await getAttachmentContent('m1', 'a1');
    expect(mGet).toHaveBeenCalledWith('/messages/m1/attachments/a1');
  });
  it('adds an attachment', async () => {
    mPost.mockResolvedValue(undefined);
    await addAttachment('m1', { name: 'n', contentType: 't', contentBytes: 'b' });
    const body = mPost.mock.calls[0][1] as any;
    expect(mPost.mock.calls[0][0]).toBe('/messages/m1/attachments');
    expect(body.Name).toBe('n');
  });
});

describe('getConversation', () => {
  it('sorts messages by received date and escapes quotes', async () => {
    mGet.mockResolvedValue({ value: [
      { Id: '2', ReceivedDateTime: '2024-02-01' },
      { Id: '1', ReceivedDateTime: '2024-01-01' },
      { Id: '3' },
    ] });
    const res = await getConversation("conv'id");
    expect(res.map(m => m.Id)).toEqual(['3', '1', '2']);
    const params = mGet.mock.calls[0][1] as Record<string, string>;
    expect(params['$filter']).toBe("ConversationId eq 'conv''id'");
    expect(params['$top']).toBe('50');
  });
  it('honours custom top', async () => {
    mGet.mockResolvedValue({ value: [] });
    await getConversation('c', 5);
    expect((mGet.mock.calls[0][1] as Record<string, string>)['$top']).toBe('5');
  });
});

describe('setCategories', () => {
  it('patches categories', async () => {
    mPatch.mockResolvedValue({ Id: 'm1' });
    await setCategories('m1', ['Red', 'Blue']);
    expect(mPatch).toHaveBeenCalledWith('/messages/m1', { Categories: ['Red', 'Blue'] });
  });
});

describe('getConversation sort with missing dates on both sides', () => {
  it('orders entries when the later one lacks a date', async () => {
    vi.mocked(owaGet).mockResolvedValue({ value: [
      { Id: 'a', ReceivedDateTime: '2024-01-01' },
      { Id: 'b' },
    ] });
    const res = await getConversation('c');
    expect(res.map(m => m.Id)).toEqual(['b', 'a']);
  });
});
