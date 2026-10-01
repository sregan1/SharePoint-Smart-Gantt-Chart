import { IDependencyLink, ITask, IWorkingCalendar } from '../models';

// Scheduling helpers. All dates are ISO 'yyyy-MM-dd' calendar days; arithmetic
// is done on UTC day numbers so it never depends on the viewer's timezone or
// DST. Working-day math honors IWorkingCalendar (weekdays + holidays).

const DAY_MS = 86400000;
const ISO_RE = /^(\d{4})-(\d{2})-(\d{2})/;

function toMs(iso: string): number {
  const m = ISO_RE.exec(iso || '');
  return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) : NaN;
}

function fromMs(ms: number): string {
  const d = new Date(ms);
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
}

interface ICalLookup { days: boolean[]; holidays: Set<string>; anyWorking: boolean; }
const calCache = new WeakMap<IWorkingCalendar, { key: string; lookup: ICalLookup }>();

function lookupFor(cal: IWorkingCalendar): ICalLookup {
  const key = `${cal.workingDays.join('')}|${cal.holidays.join(',')}`;
  const hit = calCache.get(cal);
  if (hit && hit.key === key) return hit.lookup;
  const days = [false, false, false, false, false, false, false];
  cal.workingDays.forEach(d => { if (d >= 0 && d <= 6) days[d] = true; });
  const lookup: ICalLookup = {
    days,
    holidays: new Set(cal.holidays),
    anyWorking: days.some(Boolean),
  };
  calCache.set(cal, { key, lookup });
  return lookup;
}

function isWorkingMs(ms: number, lk: ICalLookup): boolean {
  // A calendar with no working weekdays would make every walk infinite; treat
  // every day as working in that degenerate case.
  if (!lk.anyWorking) return true;
  if (!lk.days[new Date(ms).getUTCDay()]) return false;
  return !lk.holidays.has(fromMs(ms));
}

export function isWorkingDay(dateISO: string, cal: IWorkingCalendar): boolean {
  const ms = toMs(dateISO);
  return isNaN(ms) ? false : isWorkingMs(ms, lookupFor(cal));
}

/**
 * The date `n` working days after (n > 0) or before (n < 0) `dateISO`.
 * n === 0 returns the date unchanged. The start date itself is never counted.
 */
export function addWorkingDays(dateISO: string, n: number, cal: IWorkingCalendar): string {
  let ms = toMs(dateISO);
  if (isNaN(ms) || !n) return dateISO;
  const lk = lookupFor(cal);
  const step = n > 0 ? DAY_MS : -DAY_MS;
  let left = Math.abs(Math.round(n));
  while (left > 0) {
    ms += step;
    if (isWorkingMs(ms, lk)) left--;
  }
  return fromMs(ms);
}

/**
 * Working days stepped from a to b: the number of working days in (a, b] when
 * b is later, negated when b is earlier. Inverse of addWorkingDays.
 */
export function workingDaysBetween(aISO: string, bISO: string, cal: IWorkingCalendar): number {
  const a = toMs(aISO);
  const b = toMs(bISO);
  if (isNaN(a) || isNaN(b) || a === b) return 0;
  const lk = lookupFor(cal);
  const lo = Math.min(a, b);
  const hi = Math.max(a, b);
  let count = 0;
  for (let ms = lo + DAY_MS; ms <= hi; ms += DAY_MS) {
    if (isWorkingMs(ms, lk)) count++;
  }
  return a < b ? count : -count;
}

// ─── Dependency graph helpers ────────────────────────────────────────────────

const DEFAULT_LINK: IDependencyLink = { type: 'FS', lag: 0 };

function linkOf(task: ITask, predId: number): IDependencyLink {
  return (task.dependencyLinks && task.dependencyLinks[predId]) || DEFAULT_LINK;
}

/** predecessor id -> successor tasks (only edges between known tasks). */
function buildSuccessors(tasks: ITask[]): Map<number, ITask[]> {
  const known = new Set(tasks.map(t => t.id));
  const succ = new Map<number, ITask[]>();
  tasks.forEach(t => {
    (t.dependencies || []).forEach(p => {
      if (p === t.id || !known.has(p)) return;
      const list = succ.get(p);
      if (list) list.push(t); else succ.set(p, [t]);
    });
  });
  return succ;
}

/**
 * Would adding a link fromId (predecessor) -> toId (dependent) create a cycle?
 * True when fromId === toId or fromId already depends, directly or
 * transitively, on toId.
 */
export function findDependencyCycle(tasks: ITask[], fromId: number, toId: number): boolean {
  if (fromId === toId) return true;
  const succ = buildSuccessors(tasks);
  const seen = new Set<number>([toId]);
  const stack = [toId];
  while (stack.length > 0) {
    const cur = stack.pop()!;
    const next = succ.get(cur) || [];
    for (let i = 0; i < next.length; i++) {
      const id = next[i].id;
      if (id === fromId) return true;
      if (!seen.has(id)) { seen.add(id); stack.push(id); }
    }
  }
  return false;
}

// ─── Cascading shifts ────────────────────────────────────────────────────────

interface IDates { start: string; due: string; }

function laterOf(a: string, b: string): string { return a >= b ? a : b; }

/**
 * Earliest start/due a dependent may have given one predecessor's dates.
 * `start`/`due` are lower bounds (empty string = no constraint).
 */
function constraintFrom(pred: IDates, link: IDependencyLink, cal: IWorkingCalendar): { start: string; due: string } {
  switch (link.type) {
    case 'SS': return { start: addWorkingDays(pred.start, link.lag, cal), due: '' };
    case 'FF': return { start: '', due: addWorkingDays(pred.due, link.lag, cal) };
    case 'SF': return { start: '', due: addWorkingDays(pred.start, link.lag, cal) };
    default:   return { start: addWorkingDays(pred.due, 1 + link.lag, cal), due: '' };
  }
}

/**
 * After `movedTaskId` changed dates (tasks must already contain the new
 * dates), returns the new dates of every dependent that has to move so that
 * all links are satisfied again. Dependents only ever move later, keep their
 * working-day duration, and tasks caught in a dependency cycle are left alone.
 */
export function computeShifts(
  tasks: ITask[],
  movedTaskId: number,
  cal: IWorkingCalendar
): Array<{ id: number; startDate: string; dueDate: string }> {
  const succ = buildSuccessors(tasks);
  const byId = new Map<number, ITask>(tasks.map(t => [t.id, t]));
  if (!byId.has(movedTaskId)) return [];

  // Descendants of the moved task.
  const reach = new Set<number>();
  const stack = [movedTaskId];
  while (stack.length > 0) {
    const cur = stack.pop()!;
    (succ.get(cur) || []).forEach(t => {
      if (t.id !== movedTaskId && !reach.has(t.id)) { reach.add(t.id); stack.push(t.id); }
    });
  }
  if (reach.size === 0) return [];

  // Kahn ordering restricted to the reachable subgraph; cycle members never
  // reach in-degree 0 and are dropped.
  const indeg = new Map<number, number>();
  reach.forEach(id => indeg.set(id, 0));
  reach.forEach(id => {
    (succ.get(id) || []).forEach(t => {
      if (reach.has(t.id)) indeg.set(t.id, (indeg.get(t.id) || 0) + 1);
    });
  });
  const queue: number[] = [];
  // Entry points: reachable tasks fed (within the subgraph) by no one but the moved task.
  reach.forEach(id => { if ((indeg.get(id) || 0) === 0) queue.push(id); });
  const order: number[] = [];
  while (queue.length > 0) {
    const id = queue.shift()!;
    order.push(id);
    (succ.get(id) || []).forEach(t => {
      if (!reach.has(t.id)) return;
      const d = (indeg.get(t.id) || 0) - 1;
      indeg.set(t.id, d);
      if (d === 0) queue.push(t.id);
    });
  }

  const current = new Map<number, IDates>();
  tasks.forEach(t => current.set(t.id, { start: t.startDate, due: t.dueDate }));
  const changed = new Set<number>();

  order.forEach(id => {
    const task = byId.get(id)!;
    const cur = current.get(id)!;
    if (!cur.start || !cur.due) return;

    let needStart = '';
    let needDue = '';
    (task.dependencies || []).forEach(p => {
      const pd = current.get(p);
      if (!pd || !pd.start || !pd.due || !byId.has(p)) return;
      const c = constraintFrom(pd, linkOf(task, p), cal);
      if (c.start) needStart = laterOf(needStart, c.start);
      if (c.due) needDue = laterOf(needDue, c.due);
    });

    const span = Math.max(0, workingDaysBetween(cur.start, cur.due, cal));
    let newStart = cur.start;
    let newDue = cur.due;
    if (needStart && needStart > newStart) {
      newStart = needStart;
      newDue = addWorkingDays(newStart, span, cal);
    }
    if (needDue && needDue > newDue) {
      newDue = needDue;
      const s = addWorkingDays(newDue, -span, cal);
      if (s > newStart) newStart = s;
    }
    if (newDue < newStart) newDue = newStart;

    if (newStart !== cur.start || newDue !== cur.due) {
      current.set(id, { start: newStart, due: newDue });
      changed.add(id);
    }
  });

  const result: Array<{ id: number; startDate: string; dueDate: string }> = [];
  order.forEach(id => {
    if (changed.has(id)) {
      const c = current.get(id)!;
      result.push({ id, startDate: c.start, dueDate: c.due });
    }
  });
  return result;
}

// ─── Critical path ───────────────────────────────────────────────────────────

/**
 * Critical path via forward/backward pass on working-day indexes. Each task's
 * planned start is kept as a lower bound for its early start; tasks with zero
 * float against the latest finish are critical. Cancelled and undated tasks
 * are ignored; tasks inside a dependency cycle are never critical.
 */
export function computeCriticalPath(tasks: ITask[], cal: IWorkingCalendar): Set<number> {
  const result = new Set<number>();
  const nodes = tasks.filter(t => t.status !== 'Cancelled' && t.startDate && t.dueDate && !isNaN(toMs(t.startDate)) && !isNaN(toMs(t.dueDate)));
  if (nodes.length === 0) return result;

  const byId = new Map<number, ITask>(nodes.map(t => [t.id, t]));
  let origin = nodes[0].startDate;
  nodes.forEach(t => { if (t.startDate < origin) origin = t.startDate; });

  const dur = new Map<number, number>();
  const es = new Map<number, number>();
  nodes.forEach(t => {
    const s = workingDaysBetween(origin, t.startDate, cal);
    const e = workingDaysBetween(origin, t.dueDate, cal);
    es.set(t.id, s);
    dur.set(t.id, Math.max(1, e - s + 1));
  });

  const succ = buildSuccessors(nodes);
  const indeg = new Map<number, number>();
  nodes.forEach(t => indeg.set(t.id, 0));
  nodes.forEach(t => {
    (succ.get(t.id) || []).forEach(s => indeg.set(s.id, (indeg.get(s.id) || 0) + 1));
  });
  const queue: number[] = [];
  indeg.forEach((d, id) => { if (d === 0) queue.push(id); });
  const order: number[] = [];
  while (queue.length > 0) {
    const id = queue.shift()!;
    order.push(id);
    (succ.get(id) || []).forEach(s => {
      const d = (indeg.get(s.id) || 0) - 1;
      indeg.set(s.id, d);
      if (d === 0) queue.push(s.id);
    });
  }
  const inOrder = new Set(order);

  // Forward pass
  const ef = new Map<number, number>();
  order.forEach(id => {
    const t = byId.get(id)!;
    const d = dur.get(id)!;
    let s = es.get(id)!;
    (t.dependencies || []).forEach(p => {
      if (!inOrder.has(p) || !byId.has(p)) return;
      const ps = es.get(p)!;
      const pf = ef.get(p)!;
      const l = linkOf(t, p);
      if (l.type === 'FS') s = Math.max(s, pf + 1 + l.lag);
      else if (l.type === 'SS') s = Math.max(s, ps + l.lag);
      else if (l.type === 'FF') s = Math.max(s, pf + l.lag - d + 1);
      else s = Math.max(s, ps + l.lag - d + 1);
    });
    es.set(id, s);
    ef.set(id, s + d - 1);
  });

  let projectEnd = -Infinity;
  order.forEach(id => { projectEnd = Math.max(projectEnd, ef.get(id)!); });

  // Backward pass (reverse topological order)
  const lf = new Map<number, number>();
  for (let i = order.length - 1; i >= 0; i--) {
    const id = order[i];
    const d = dur.get(id)!;
    let f = projectEnd;
    (succ.get(id) || []).forEach(t => {
      if (!inOrder.has(t.id)) return;
      const l = linkOf(t, id);
      const jf = lf.get(t.id)!;
      const js = jf - dur.get(t.id)! + 1;
      if (l.type === 'FS') f = Math.min(f, js - 1 - l.lag);
      else if (l.type === 'SS') f = Math.min(f, js - l.lag + d - 1);
      else if (l.type === 'FF') f = Math.min(f, jf - l.lag);
      else f = Math.min(f, jf - l.lag + d - 1);
    });
    lf.set(id, f);
    if (f - ef.get(id)! <= 0) result.add(id);
  }
  return result;
}
