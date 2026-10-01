import { differenceInCalendarDays } from 'date-fns';
import {
  ITask, IDependencyLink, IWorkingCalendar, ZoomLevel, STATUS_COLORS, PRIORITY_COLORS,
  IGanttDisplaySettings, phaseColor,
} from '../../models';
import { computeTaskHealth, healthColor } from '../../utils/healthUtils';
import { parseDateOnly, dateToDateOnlyString } from '../../utils/dateUtils';
import { addWorkingDays, isWorkingDay } from '../../utils/scheduleUtils';

// ─── Layout constants ────────────────────────────────────────────────────────

export const HEADER_HEIGHT = 56;
export const BAR_HEIGHT = 26;
export const MILESTONE_SIZE = 12;
export const MIN_BAR_WIDTH = 6;
export const BUFFER_DAYS = 30;
/** Rows rendered above/below the viewport so fast scrolling doesn't flash blank. */
export const ROW_OVERSCAN = 8;

export const DAY_WIDTH: Record<ZoomLevel, number> = {
  day: 42,
  week: 18,
  month: 7,
  quarter: 4,
};

/** Critical / violation accent. Theme-aware where the host exposes the variable. */
export const CRITICAL_COLOR = 'var(--sg-critical, #D13438)';

/** Zoom level whose header layout best matches an arbitrary (fit / Ctrl+wheel) day width. */
export function zoomForDayWidth(dw: number): ZoomLevel {
  if (dw >= 30) return 'day';
  if (dw >= 11) return 'week';
  if (dw >= 5.5) return 'month';
  return 'quarter';
}

// ─── Colors ──────────────────────────────────────────────────────────────────

const HEX_COLOR_RE = /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/;

export function getTaskColor(task: ITask, settings: IGanttDisplaySettings): string {
  // task.color is free-form user input (models/index.ts) — a stray value
  // (short hex, a CSS name, import garbage) must not flow into SVG color
  // attributes unvalidated, or the bar renders with NaN-derived colors.
  if (task.color && HEX_COLOR_RE.test(task.color)) return task.color;
  if (settings.colorBy === 'priority') return PRIORITY_COLORS[task.priority] || '#0078D4';
  if (settings.colorBy === 'phase' && task.phase) return phaseColor(task.phase);
  if (settings.colorBy === 'health') return healthColor(computeTaskHealth(task));
  return STATUS_COLORS[task.status] || '#0078D4';
}

export function hexToRgba(hex: string, alpha: number): string {
  if (!HEX_COLOR_RE.test(hex)) hex = '#0078D4';
  if (hex.length === 4) {
    hex = `#${hex[1]}${hex[1]}${hex[2]}${hex[2]}${hex[3]}${hex[3]}`;
  }
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  return `rgba(${r},${g},${b},${alpha})`;
}

// SVG ids can't contain '#'; gradients are keyed by color so distinct tasks
// sharing a color share one <linearGradient> def instead of emitting one per bar.
export function colorId(hex: string): string {
  return hex.replace('#', '');
}

// ─── Dates ───────────────────────────────────────────────────────────────────

export function taskDuration(task: ITask): number {
  const s = parseDateOnly(task.startDate);
  const e = parseDateOnly(task.dueDate);
  if (!s || !e) return 0;
  return Math.max(1, differenceInCalendarDays(e, s) + 1);
}

/**
 * Snap a date onto a working day, walking in `dir` (1 = later, -1 = earlier).
 * Bounded so a degenerate calendar can't spin.
 */
export function snapToWorkingDay(d: Date, dir: 1 | -1, cal: IWorkingCalendar): Date {
  let cur = d;
  for (let i = 0; i < 400; i++) {
    if (isWorkingDay(dateToDateOnlyString(cur), cal)) return cur;
    cur = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate() + dir);
  }
  return d;
}

/** Shift an ISO date by whole working days. */
export function shiftWorkingDays(iso: string, n: number, cal: IWorkingCalendar): string {
  return addWorkingDays(iso, n, cal);
}

// ─── Dependencies ────────────────────────────────────────────────────────────

const DEFAULT_LINK: IDependencyLink = { type: 'FS', lag: 0 };

export function linkOf(task: ITask, predId: number): IDependencyLink {
  return (task.dependencyLinks && task.dependencyLinks[predId]) || DEFAULT_LINK;
}

export interface IDateRange { start: Date; end: Date }

/**
 * True when the successor's dates break the predecessor link (FS/SS/FF/SF with
 * lag in working days). Uses the same lower bounds as computeShifts.
 */
export function isLinkViolated(
  pred: IDateRange,
  succ: IDateRange,
  link: IDependencyLink,
  cal: IWorkingCalendar
): boolean {
  const ps = dateToDateOnlyString(pred.start);
  const pe = dateToDateOnlyString(pred.end);
  const ss = dateToDateOnlyString(succ.start);
  const se = dateToDateOnlyString(succ.end);
  switch (link.type) {
    case 'SS': return ss < addWorkingDays(ps, link.lag, cal);
    case 'FF': return se < addWorkingDays(pe, link.lag, cal);
    case 'SF': return se < addWorkingDays(ps, link.lag, cal);
    default:   return ss < addWorkingDays(pe, 1 + link.lag, cal);
  }
}

// ─── Row model ───────────────────────────────────────────────────────────────

export interface IVisibleRow {
  type: 'task' | 'phase';
  task?: ITask;
  phase?: string;
  isChild?: boolean;
  depth?: number;
  /** Task sits in a parent/sub-task cycle and was promoted to the top level. */
  cyclic?: boolean;
}

export function buildVisibleRows(tasks: ITask[], collapsedPhases: Set<string>): IVisibleRow[] {
  const rows: IVisibleRow[] = [];
  const ids = new Set(tasks.map(t => t.id));
  // A task only renders as a sub-task if its parent is actually present;
  // otherwise (parent deleted or filtered out) it's promoted to top level so
  // it never silently disappears.
  const isSubTask = (t: ITask): boolean => !!t.parentTaskId && t.parentTaskId !== t.id && ids.has(t.parentTaskId);

  const byPhase = new Map<string, ITask[]>();
  const noPhase: ITask[] = [];

  tasks.forEach(t => {
    if (isSubTask(t)) return; // appended under parent below
    if (t.phase) {
      if (!byPhase.has(t.phase)) byPhase.set(t.phase, []);
      byPhase.get(t.phase)!.push(t);
    } else {
      noPhase.push(t);
    }
  });

  const subtaskMap = new Map<number, ITask[]>();
  tasks.filter(isSubTask).forEach(t => {
    if (!subtaskMap.has(t.parentTaskId!)) subtaskMap.set(t.parentTaskId!, []);
    subtaskMap.get(t.parentTaskId!)!.push(t);
  });

  // Tasks in a parent cycle (A→B→A) are sub-tasks of each other, so no chain
  // of roots ever reaches them and they'd vanish. Find everything reachable
  // from a real root; the rest are cyclic and get rendered at the top level.
  const reachable = new Set<number>();
  const mark = (t: ITask): void => {
    if (reachable.has(t.id)) return;
    reachable.add(t.id);
    (subtaskMap.get(t.id) || []).forEach(mark);
  };
  tasks.forEach(t => { if (!isSubTask(t)) mark(t); });
  const cyclicTasks = tasks.filter(t => !reachable.has(t.id));

  // Recurse through every nesting level (a sub-task can itself have
  // sub-tasks — e.g. via a direct list edit or import, even though the Task
  // panel's parent picker prevents creating it through the UI) so deeper
  // descendants are flattened into view instead of silently vanishing.
  // `emitted` guards against a parent-cycle in the raw data looping forever.
  const emitted = new Set<number>();
  const addTask = (task: ITask, depth: number, cyclic: boolean): void => {
    if (emitted.has(task.id)) return;
    emitted.add(task.id);
    rows.push({ type: 'task', task, isChild: depth > 0, depth, cyclic });
    (subtaskMap.get(task.id) || []).forEach(c => addTask(c, depth + 1, cyclic));
  };

  byPhase.forEach((phaseTasks, phase) => {
    rows.push({ type: 'phase', phase });
    if (!collapsedPhases.has(phase)) {
      phaseTasks.forEach(t => addTask(t, 0, false));
    } else {
      // Keep `emitted` honest for hidden tasks so they're not re-added below.
      phaseTasks.forEach(t => markEmitted(t, subtaskMap, emitted));
    }
  });

  noPhase.forEach(t => addTask(t, 0, false));
  cyclicTasks.forEach(t => addTask(t, 0, true));

  return rows;
}

function markEmitted(task: ITask, subtaskMap: Map<number, ITask[]>, emitted: Set<number>): void {
  if (emitted.has(task.id)) return;
  emitted.add(task.id);
  (subtaskMap.get(task.id) || []).forEach(c => markEmitted(c, subtaskMap, emitted));
}

// ─── Shared option types ─────────────────────────────────────────────────────

/** Extra behavior a caller can request when reporting a task date change. */
export interface ITaskUpdateOptions {
  /** After saving, move dependents of the updated task so links stay satisfied. */
  cascade?: boolean;
  /** Cascade from this predecessor instead (used when a new link is added). */
  cascadeFrom?: number;
}

const COLLAPSED_KEY_PREFIX = 'SmartGantt_collapsed_';

/** Collapsed phases are remembered per project list; storage failures are non-fatal. */
export function loadCollapsedPhases(listName: string): Set<string> {
  try {
    const raw = window.localStorage.getItem(COLLAPSED_KEY_PREFIX + listName);
    const arr = raw ? JSON.parse(raw) : [];
    return new Set<string>(Array.isArray(arr) ? arr.filter(x => typeof x === 'string') : []);
  } catch {
    return new Set<string>();
  }
}

export function saveCollapsedPhases(listName: string, phases: Set<string>): void {
  try {
    window.localStorage.setItem(COLLAPSED_KEY_PREFIX + listName, JSON.stringify(Array.from(phases)));
  } catch {
    // Storage unavailable — collapse state just won't persist.
  }
}
