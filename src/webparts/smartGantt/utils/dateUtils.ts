import { format } from 'date-fns';

// Schedule dates (task/project start and due) are calendar days, not instants.
// SharePoint stores them as UTC-midnight DateTimes; the service layer normalizes
// everything to 'YYYY-MM-DD' strings, and components parse those as *local*
// dates so the displayed day never shifts with the viewer's timezone.
// Created/Modified timestamps are real instants and must NOT go through these
// helpers — parse those with `new Date()` directly.

const DATE_ONLY_RE = /^(\d{4})-(\d{2})-(\d{2})$/;

export function toDateOnly(value: string | null | undefined): string {
  if (!value) return '';
  if (DATE_ONLY_RE.test(value)) return value;
  const d = new Date(value);
  if (isNaN(d.getTime())) return '';
  // Read the UTC parts: current data is written as UTC midnight. Legacy values
  // written as local midnight land within ±12h of the intended day, so rolling
  // forward when the UTC time is past noon recovers the intended calendar day
  // for writers on either side of UTC.
  let y = d.getUTCFullYear();
  let m = d.getUTCMonth();
  let day = d.getUTCDate();
  if (d.getUTCHours() >= 12) {
    const rolled = new Date(Date.UTC(y, m, day + 1));
    y = rolled.getUTCFullYear();
    m = rolled.getUTCMonth();
    day = rolled.getUTCDate();
  }
  return `${y}-${String(m + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

/** Parse a schedule date to local midnight. Returns null for empty/invalid input. */
export function parseDateOnly(value: string | null | undefined): Date | null {
  const s = toDateOnly(value);
  const m = DATE_ONLY_RE.exec(s);
  if (!m) return null;
  return new Date(+m[1], +m[2] - 1, +m[3]);
}

/** Format a schedule date string for display; falls back when empty/invalid. */
export function formatDateOnly(value: string | null | undefined, fmt: string, fallback = '—'): string {
  const d = parseDateOnly(value);
  return d ? format(d, fmt) : fallback;
}

/** Convert a local Date back to the canonical 'YYYY-MM-DD' form. */
export function dateToDateOnlyString(d: Date): string {
  return format(d, 'yyyy-MM-dd');
}

/** Local midnight of today — the reference point for all overdue comparisons. */
export function todayLocalMidnight(): Date {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}
