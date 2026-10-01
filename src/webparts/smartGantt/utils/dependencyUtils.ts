import { DependencyType, IDependencyLink } from '../models';

// Dependencies are stored in a single Text column as a comma-separated list of
// predecessor ids, each optionally followed by a link type and lag in working
// days: "12,15SS+2,18FF-1". A bare id means finish-to-start with no lag, so
// data written before link types existed reads back unchanged.

// Lag units: d/ed/days (working days), w/wk/weeks (= 5 working days).
const TOKEN_RE = /^(\d+)\s*(?:(FS|SS|FF|SF))?\s*(?:([+-])\s*(\d+)\s*(d|ed|days?|w|wks?|weeks?)?)?$/i;

export interface IParsedDependencyToken {
  id: number;
  link: IDependencyLink;
}

/** Parse one token such as "15", "15SS", "15SS+2", "15FF-1d", "3FS+1w". Returns null if unparseable. */
export function parseDependencyToken(token: string): IParsedDependencyToken | null {
  const m = TOKEN_RE.exec((token || '').trim());
  if (!m) return null;
  const id = parseInt(m[1], 10);
  if (!id || id <= 0) return null;
  const type = ((m[2] || 'FS').toUpperCase()) as DependencyType;
  let lag = m[4] ? parseInt(m[4], 10) : 0;
  if (m[5] && m[5].toLowerCase().charAt(0) === 'w') lag *= 5;
  if (m[3] === '-') lag = -lag;
  return { id, link: { type, lag } };
}

export function isDefaultLink(link: IDependencyLink | undefined): boolean {
  return !link || (link.type === 'FS' && !link.lag);
}

export function parseDependencies(raw: string): { ids: number[]; links: Record<number, IDependencyLink> } {
  const ids: number[] = [];
  const links: Record<number, IDependencyLink> = {};
  (raw || '').split(/[,;]/).forEach(part => {
    const parsed = parseDependencyToken(part);
    if (!parsed || ids.indexOf(parsed.id) !== -1) return;
    ids.push(parsed.id);
    // Only non-default links are kept; a missing key means FS, lag 0.
    if (!isDefaultLink(parsed.link)) links[parsed.id] = parsed.link;
  });
  return { ids, links };
}

/** Short label, e.g. "FS", "SS+2", "FF-1". */
export function formatDependencyLink(link: IDependencyLink | undefined): string {
  const type = link?.type || 'FS';
  const lag = link?.lag || 0;
  return lag === 0 ? type : `${type}${lag > 0 ? '+' : '-'}${Math.abs(lag)}`;
}

export function serializeDependencies(ids: number[], links?: Record<number, IDependencyLink>): string {
  return (ids || [])
    .map(id => {
      const link = links ? links[id] : undefined;
      return isDefaultLink(link) ? String(id) : `${id}${formatDependencyLink(link)}`;
    })
    .join(',');
}

/** Lag/type suffix of a token whose leading part is not a numeric id, e.g. the "FS+2d" of "Design Review FS+2d". */
const SUFFIX_RE = /^(.*?\S)\s+(?:(FS|SS|FF|SF)\s*)?(?:([+-])\s*(\d+)\s*(d|ed|days?|w|wks?|weeks?)?)?$/i;

/** Split "<name> [FS|SS|FF|SF][+/-N[d|w]]" into the name and its link; returns null when there is no suffix. */
export function splitNamedLink(token: string): { name: string; link: IDependencyLink } | null {
  const m = SUFFIX_RE.exec((token || '').trim());
  if (!m || (!m[2] && !m[3])) return null;
  let lag = m[4] ? parseInt(m[4], 10) : 0;
  if (m[5] && m[5].toLowerCase().charAt(0) === 'w') lag *= 5;
  if (m[3] === '-') lag = -lag;
  return { name: m[1].trim(), link: { type: ((m[2] || 'FS').toUpperCase()) as DependencyType, lag } };
}
