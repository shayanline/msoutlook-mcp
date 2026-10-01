import { beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('./client.js', () => ({ owaCloudSettingsGet: vi.fn() }));
vi.mock('../utils/logger.js', () => ({ logger: { warn: vi.fn() } }));

import { owaCloudSettingsGet } from './client.js';
import { applyOutlookSignature } from './signatures.js';
import { logger } from '../utils/logger.js';

const mCloudGet = vi.mocked(owaCloudSettingsGet);

beforeEach(() => vi.clearAllMocks());

describe('applyOutlookSignature', () => {
  it('appends the selected new HTML signature unchanged', async () => {
    const signature = '<div><img src="data:image/png;base64,AAA"><a href="https://example.com/book?a=1&amp;b=2">Book</a></div>';
    mCloudGet
      .mockResolvedValueOnce([{ name: 'roaming_new_signature', value: 'Work' }])
      .mockResolvedValueOnce([{ name: 'Work', secondaryKey: 'htm', value: signature }]);

    const result = await applyOutlookSignature('Hello\nworld', 'new', 'HTML');

    expect(result).toBe(`Hello<br>world<br><br>${signature}`);
    expect(mCloudGet).toHaveBeenNthCalledWith(1, '/settings/', {
      settingname: 'signaturehtml,signaturetxt,roaming_new_signature,roaming_reply_signature',
    });
    expect(mCloudGet).toHaveBeenNthCalledWith(2, '/settings/account', { settingname: 'Work' }, true);
  });

  it('uses the selected reply text signature for replies and forwards', async () => {
    mCloudGet
      .mockResolvedValueOnce([{ name: 'roaming_reply_signature', value: 'Short' }])
      .mockResolvedValueOnce([{ name: 'Short', secondaryKey: 'txt', value: 'Regards\nSam' }]);

    await expect(applyOutlookSignature('Reply body', 'reply', 'Text')).resolves.toBe(
      'Reply body\n\nRegards\nSam',
    );
  });

  it('uses the legacy Outlook signature when roaming signatures are unavailable', async () => {
    mCloudGet.mockResolvedValueOnce([
      { name: 'signaturehtml', value: '<div>Legacy signature</div>' },
      { name: 'signaturetxt', value: 'Legacy signature' },
    ]);

    await expect(applyOutlookSignature('Hello', 'new', 'HTML')).resolves.toBe(
      'Hello<br><br><div>Legacy signature</div>',
    );
    expect(mCloudGet).toHaveBeenCalledTimes(1);
  });

  it('returns the formatted body when no default signature is selected', async () => {
    mCloudGet.mockResolvedValueOnce([]);

    await expect(applyOutlookSignature('Hello\nworld', 'new', 'HTML')).resolves.toBe('Hello<br>world');
    expect(mCloudGet).toHaveBeenCalledTimes(1);
  });

  it('uses the signature without a leading gap when the body is empty', async () => {
    const signature = '<div>Sam</div>';
    mCloudGet
      .mockResolvedValueOnce([{ name: 'roaming_reply_signature', value: 'Short' }])
      .mockResolvedValueOnce([{ name: 'Short', secondaryKey: 'htm', value: signature }]);

    await expect(applyOutlookSignature('', 'reply', 'HTML')).resolves.toBe(signature);
  });

  it('does not append a signature already present in the body', async () => {
    const signature = '<div>Sam</div>';
    mCloudGet
      .mockResolvedValueOnce([{ name: 'roaming_new_signature', value: 'Work' }])
      .mockResolvedValueOnce([{ name: 'Work', secondaryKey: 'htm', value: signature }]);

    await expect(applyOutlookSignature(`Hello<br>${signature}`, 'new', 'HTML')).resolves.toBe(
      `Hello<br>${signature}`,
    );
  });

  it('keeps email composition available when cloud settings fail', async () => {
    mCloudGet.mockRejectedValueOnce(new Error('endpoint changed'));

    await expect(applyOutlookSignature('Hello', 'new', 'HTML')).resolves.toBe('Hello');
    expect(logger.warn).toHaveBeenCalledWith('Could not load Outlook signature', 'endpoint changed');
  });
});
