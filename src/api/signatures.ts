import { owaCloudSettingsGet } from './client.js';
import { toHtmlBody } from './mail.js';
import { logger } from '../utils/logger.js';

export type SignatureKind = 'new' | 'reply';
export type SignatureFormat = 'HTML' | 'Text';

interface CloudSetting {
  name: string;
  value?: string;
  secondaryKey?: string;
}

export async function applyOutlookSignature(
  body: string,
  kind: SignatureKind,
  format: SignatureFormat,
): Promise<string> {
  const content = format === 'HTML' ? toHtmlBody(body) : body;

  try {
    const defaults = await owaCloudSettingsGet<CloudSetting[]>('/settings/', {
      settingname: 'signaturehtml,signaturetxt,roaming_new_signature,roaming_reply_signature',
    });
    const defaultName = defaults.find(
      setting => setting.name === `roaming_${kind}_signature`,
    )?.value?.trim();
    const formatKey = format === 'HTML' ? 'htm' : 'txt';
    const legacyKey = format === 'HTML' ? 'signaturehtml' : 'signaturetxt';
    let signature = defaults.find(setting => setting.name === legacyKey)?.value;

    if (defaultName) {
      const signatures = await owaCloudSettingsGet<CloudSetting[]>(
        '/settings/account',
        { settingname: defaultName },
        true,
      );
      signature = signatures.find(setting => setting.secondaryKey === formatKey)?.value;
    }
    if (!signature || content.includes(signature)) return content;
    if (!content) return signature;

    return `${content}${format === 'HTML' ? '<br><br>' : '\n\n'}${signature}`;
  } catch (error) {
    logger.warn('Could not load Outlook signature', error instanceof Error ? error.message : String(error));
    return content;
  }
}
