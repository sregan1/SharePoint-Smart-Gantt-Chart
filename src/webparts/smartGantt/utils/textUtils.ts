const NAMED_ENTITIES: Record<string, string> = {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ',
};

/**
 * Decodes HTML character references (`&#58;`, `&#x3a;`, `&amp;`, ...) that
 * SharePoint's rich-text editor writes into multiline text columns, e.g.
 * `https&#58;//host/path` back to `https://host/path`. Plain text without
 * entities is returned unchanged.
 */
export function decodeHtmlEntities(text: string): string {
  if (!text || text.indexOf('&') === -1) return text;
  return text.replace(/&(?:#(\d+)|#[xX]([0-9a-fA-F]+)|([a-zA-Z]+));/g, (whole, dec, hex, name) => {
    if (dec) return String.fromCharCode(parseInt(dec, 10));
    if (hex) return String.fromCharCode(parseInt(hex, 16));
    return NAMED_ENTITIES[name] !== undefined ? NAMED_ENTITIES[name] : whole;
  });
}
