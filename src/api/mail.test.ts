import { describe, it, expect } from 'vitest';
import { looksLikeHtml, toHtmlBody } from './mail.js';

describe('looksLikeHtml', () => {
  it('detects real html tags', () => {
    expect(looksLikeHtml('<br>')).toBe(true);
    expect(looksLikeHtml('<div>hi</div>')).toBe(true);
    expect(looksLikeHtml('<ul><li>a</li></ul>')).toBe(true);
  });

  it('treats plain text and lone angle brackets as not html', () => {
    expect(looksLikeHtml('Integrations > Vincere')).toBe(false);
    expect(looksLikeHtml('a < b and c > d')).toBe(false);
    expect(looksLikeHtml('just text')).toBe(false);
  });
});

describe('toHtmlBody', () => {
  it('passes through content that already contains html', () => {
    const html = '<div>Hi</div><br><ul><li>x</li></ul>';
    expect(toHtmlBody(html)).toBe(html);
  });

  it('converts single newlines to <br>', () => {
    expect(toHtmlBody('one\ntwo')).toBe('one<br>two');
  });

  it('converts a blank line to a paragraph gap', () => {
    expect(toHtmlBody('Hi,\n\nGood news.')).toBe('Hi,<br><br>Good news.');
  });

  it('normalises CRLF and CR before converting', () => {
    expect(toHtmlBody('one\r\ntwo')).toBe('one<br>two');
    expect(toHtmlBody('one\rtwo')).toBe('one<br>two');
  });

  it('escapes html-significant characters in plain text', () => {
    expect(toHtmlBody('a < b & c > d')).toBe('a &lt; b &amp; c &gt; d');
  });
});
