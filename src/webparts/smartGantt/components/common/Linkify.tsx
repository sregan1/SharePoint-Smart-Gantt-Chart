import * as React from 'react';
import { decodeHtmlEntities } from '../../utils/textUtils';

const URL_PATTERN = /https?:\/\/[^\s<>"']+/gi;
const TRAILING_PUNCTUATION = /[.,;:!?)\]}]+$/;

/** True when the text contains at least one http(s) URL. */
export function containsUrl(text: string | undefined | null): boolean {
  if (!text) return false;
  return new RegExp(URL_PATTERN.source, 'i').test(decodeHtmlEntities(text));
}

/**
 * Renders text with http(s) URLs turned into links that open in a new tab.
 * Builds React elements only (no HTML injection). HTML entities are decoded
 * first so text that SharePoint rich-text-encoded (`https&#58;//`) still
 * yields a working URL.
 */
export const Linkify: React.FC<{ text: string }> = ({ text }) => {
  const source = decodeHtmlEntities(text);
  const parts: React.ReactNode[] = [];
  let last = 0;
  let match: RegExpExecArray | null;
  const re = new RegExp(URL_PATTERN.source, 'gi');
  while ((match = re.exec(source)) !== null) {
    const url = match[0].replace(TRAILING_PUNCTUATION, '');
    if (match.index > last) parts.push(source.slice(last, match.index));
    parts.push(
      <a key={match.index} href={url} target="_blank" rel="noopener noreferrer"
        style={{ wordBreak: 'break-all' }}>
        {url}
      </a>,
    );
    last = match.index + url.length;
    re.lastIndex = last;
  }
  if (last < source.length) parts.push(source.slice(last));
  return <span style={{ whiteSpace: 'pre-wrap' }}>{parts}</span>;
};
