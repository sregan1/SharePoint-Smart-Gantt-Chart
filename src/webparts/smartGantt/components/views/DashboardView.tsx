import * as React from 'react';
import { addDays, differenceInCalendarDays } from 'date-fns';
import * as strings from 'SmartGanttWebPartStrings';
import { formatString } from '../localeUtils';
import {
  IProject, ITask, TaskStatus, TaskPriority, IWorkingCalendar, DEFAULT_WORKING_CALENDAR,
  STATUS_COLORS, STATUS_LIGHT_COLORS, PRIORITY_COLORS, phaseColor,
  PROJECT_STATUS_COLORS, PROJECT_STATUS_LIGHT_COLORS,
} from '../../models';
import { parseDateOnly, formatDateOnly, todayLocalMidnight, dateToDateOnlyString } from '../../utils/dateUtils';
import { computeCriticalPath } from '../../utils/scheduleUtils';
import { getStatusLabel, getPriorityLabel, getProjectStatusLabel, initials, stringToColor } from '../../utils/taskDisplayUtils';
import { onActivate } from '../common/a11y';
import { THEME as T, SHARED_VIEW_CSS } from '../common/theme';

interface IDashboardViewProps {
  project: IProject;
  tasks: ITask[];
  onEditTask: (task: ITask) => void;
  onAddTask: () => void;
  /** Working calendar used for the critical-path list. Defaults to Mon–Fri. */
  calendar?: IWorkingCalendar;
}

function fmtDate(s: string): string {
  return formatDateOnly(s, 'MMM d');
}

// Modified/Created are real timestamps (not schedule dates) — parse directly.
function parseTimestamp(s: string): Date | null {
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function localMidnight(d: Date): Date {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

// ── Stat card ──────────────────────────────────────────────────────────────
const StatCard: React.FC<{
  label: string; value: number; total: number;
  color: string; bg: string; onClick?: () => void;
}> = ({ label, value, total, color, bg, onClick }) => (
  <div
    className={`sgCard${onClick ? ' sgFocusable' : ''}`}
    onClick={onClick}
    role={onClick ? 'button' : undefined}
    tabIndex={onClick ? 0 : undefined}
    aria-label={onClick ? `${label}: ${value}` : undefined}
    onKeyDown={onClick ? onActivate(onClick) : undefined}
    style={{
      flex: 1, minWidth: 110,
      background: bg, border: `1px solid ${color}30`,
      borderRadius: 8, padding: '14px 16px',
      cursor: onClick ? 'pointer' : 'default',
      transition: 'box-shadow 0.15s',
    }}
    onMouseEnter={e => { if (onClick) (e.currentTarget as HTMLElement).style.boxShadow = `0 2px 8px ${color}30`; }}
    onMouseLeave={e => { (e.currentTarget as HTMLElement).style.boxShadow = 'none'; }}
  >
    <div style={{ fontSize: 30, fontWeight: 700, color, lineHeight: 1 }}>{value}</div>
    <div style={{ fontSize: 11, color: T.textSecondary, marginTop: 4 }}>{label}</div>
    {total > 0 && (
      <div style={{ marginTop: 8, height: 4, background: `${color}25`, borderRadius: 2, overflow: 'hidden' }}>
        <div className="sgFill" style={{ height: '100%', width: `${Math.round(value / total * 100)}%`, background: color, borderRadius: 2 }} />
      </div>
    )}
  </div>
);

// ── Section header ─────────────────────────────────────────────────────────
const SectionHeader: React.FC<{ title: string }> = ({ title }) => (
  <div style={{
    fontSize: 11, fontWeight: 700, color: T.textSecondary,
    letterSpacing: '0.6px', textTransform: 'uppercase',
    paddingBottom: 8, borderBottom: `1px solid ${T.border}`, marginBottom: 14,
  }}>
    {title}
  </div>
);

// ── Task row ───────────────────────────────────────────────────────────────
const TaskRow: React.FC<{ task: ITask; onClick: () => void; showDue?: boolean; isOverdue?: boolean }> = ({
  task, onClick, showDue, isOverdue,
}) => {
  const sc = STATUS_COLORS[task.status as TaskStatus] || '#8B929A';
  return (
    <div
      className="sgFocusable"
      onClick={onClick}
      role="button"
      tabIndex={0}
      aria-label={formatString(strings.View_Dashboard_TaskRowAriaLabel, { taskTitle: task.title, status: getStatusLabel(task.status as TaskStatus) })}
      onKeyDown={onActivate(onClick)}
      style={{
        display: 'flex', alignItems: 'center', gap: 10,
        padding: '7px 10px', borderRadius: 5, cursor: 'pointer',
        background: isOverdue ? '#FDF3F4' : 'transparent',
        border: isOverdue ? '1px solid #F9D8D8' : '1px solid transparent',
        marginBottom: 4, transition: 'background 0.1s',
      }}
      onMouseEnter={e => { (e.currentTarget as HTMLElement).style.background = isOverdue ? '#FAE5E5' : 'var(--neutralLighter, #F3F2F1)'; }}
      onMouseLeave={e => { (e.currentTarget as HTMLElement).style.background = isOverdue ? '#FDF3F4' : 'transparent'; }}
    >
      <div className="sgFill" style={{ width: 8, height: 8, borderRadius: '50%', background: sc, flexShrink: 0 }} />
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontSize: 13, color: isOverdue ? '#323130' : T.text, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
          {task.isMilestone ? '◆ ' : ''}{task.title}
        </div>
        <div style={{ fontSize: 11, color: isOverdue ? '#605E5C' : T.textSecondary, marginTop: 1 }}>
          <span style={{ color: sc }}>{getStatusLabel(task.status as TaskStatus)}</span>
          {task.percentComplete > 0 && <span>  •  {task.percentComplete}%</span>}
          {task.phase && <span>  •  {task.phase}</span>}
        </div>
      </div>
      {showDue && (
        <div style={{ fontSize: 11, fontWeight: 600, color: isOverdue ? '#D13438' : T.textSecondary, flexShrink: 0 }}>
          {isOverdue ? '⚠ ' : ''}{fmtDate(task.dueDate)}
        </div>
      )}
      {task.percentComplete > 0 && !showDue && (
        <div style={{ width: 60, flexShrink: 0 }}>
          <div style={{ height: 4, background: T.border, borderRadius: 2, overflow: 'hidden' }}>
            <div className="sgFill" style={{ height: '100%', width: `${task.percentComplete}%`, background: sc, borderRadius: 2 }} />
          </div>
        </div>
      )}
    </div>
  );
};

// ── Burndown ───────────────────────────────────────────────────────────────
interface IBurndownData {
  start: Date;
  end: Date;
  total: number;
  /** Remaining-task count over time, from the start up to today (capped at the end). */
  actual: Array<{ date: Date; remaining: number }>;
  today: Date;
}

const MAX_BURNDOWN_SAMPLES = 60;

// Remaining-tasks burndown. The list has no "completed on" column, so a
// Completed task's `modified` date is used as its completion date (the last
// edit of a finished task is almost always the one that finished it).
function buildBurndown(scope: ITask[], project: IProject, today: Date): IBurndownData | null {
  if (scope.length === 0) return null;

  let start = parseDateOnly(project.startDate);
  let end = parseDateOnly(project.dueDate);
  scope.forEach(t => {
    const s = parseDateOnly(t.startDate);
    const d = parseDateOnly(t.dueDate);
    if (s && (!start || s < start)) start = s;
    if (d && (!end || d > end)) end = d;
  });
  if (!start || !end || end <= start) return null;

  const startMs = start.getTime();
  const completions: number[] = scope
    .filter(t => t.status === 'Completed')
    .map(t => {
      const m = parseTimestamp(t.modified);
      return Math.max(startMs, (m ? localMidnight(m) : today).getTime());
    });

  const last = today < end ? today : end;
  const actual: Array<{ date: Date; remaining: number }> = [];
  if (last >= start) {
    const days = differenceInCalendarDays(last, start);
    const step = Math.max(1, Math.ceil(days / MAX_BURNDOWN_SAMPLES));
    const sample = (d: Date): void => {
      const ms = d.getTime();
      actual.push({ date: d, remaining: scope.length - completions.filter(c => c <= ms).length });
    };
    for (let i = 0; i < days; i += step) sample(addDays(start, i));
    sample(last);
  }
  return { start, end, total: scope.length, actual, today };
}

const CH_W = 640;
const CH_H = 240;
const CH_L = 40;
const CH_R = 16;
const CH_T = 12;
const CH_B = 28;

const BurndownChart: React.FC<{ data: IBurndownData; color: string }> = ({ data, color }) => {
  const { start, end, total, actual, today } = data;
  const spanMs = end.getTime() - start.getTime();
  const x = (d: Date): number => CH_L + ((d.getTime() - start.getTime()) / spanMs) * (CH_W - CH_L - CH_R);
  const y = (v: number): number => CH_T + (1 - v / total) * (CH_H - CH_T - CH_B);

  const remainingNow = actual.length > 0 ? actual[actual.length - 1].remaining : total;
  const todayInRange = today >= start && today <= end;
  const actualPoints = actual.map(p => `${x(p.date).toFixed(1)},${y(p.remaining).toFixed(1)}`).join(' ');
  const summary = formatString(strings.View_Dashboard_BurndownAria, { remaining: remainingNow, total });
  const mid = Math.round(total / 2);

  return (
    <div>
      <svg
        viewBox={`0 0 ${CH_W} ${CH_H}`}
        width="100%"
        role="img"
        aria-label={summary}
        style={{ display: 'block', maxHeight: 260 }}
      >
        <title>{summary}</title>
        {/* Y gridlines + labels */}
        {[0, mid, total].filter((v, i, arr) => arr.indexOf(v) === i).map(v => (
          <g key={v}>
            <line x1={CH_L} x2={CH_W - CH_R} y1={y(v)} y2={y(v)} style={{ stroke: 'var(--neutralLight, #EDEBE9)' }} strokeWidth={1} />
            <text x={CH_L - 6} y={y(v) + 4} textAnchor="end" fontSize={10} style={{ fill: 'var(--neutralSecondary, #605E5C)' }}>{v}</text>
          </g>
        ))}
        {/* X labels */}
        <text x={CH_L} y={CH_H - 8} fontSize={10} textAnchor="start" style={{ fill: 'var(--neutralSecondary, #605E5C)' }}>{formatDateOnly(dateToDateOnlyString(start), 'MMM d')}</text>
        <text x={CH_W - CH_R} y={CH_H - 8} fontSize={10} textAnchor="end" style={{ fill: 'var(--neutralSecondary, #605E5C)' }}>{formatDateOnly(dateToDateOnlyString(end), 'MMM d')}</text>
        {/* Today marker */}
        {todayInRange && (
          <line x1={x(today)} x2={x(today)} y1={CH_T} y2={CH_H - CH_B} stroke="#D13438" strokeWidth={1} strokeDasharray="2 3" />
        )}
        {/* Ideal line */}
        <line
          x1={x(start)} y1={y(total)} x2={x(end)} y2={y(0)}
          strokeWidth={2} strokeDasharray="6 4" style={{ stroke: 'var(--neutralTertiary, #8A8886)' }}
        />
        {/* Actual line */}
        {actual.length > 1 && (
          <polyline points={actualPoints} fill="none" stroke={color} strokeWidth={2.5} strokeLinejoin="round" />
        )}
        {actual.length > 0 && (
          <circle cx={x(actual[actual.length - 1].date)} cy={y(remainingNow)} r={4} fill={color} />
        )}
      </svg>
      <div style={{ display: 'flex', gap: 16, fontSize: 11, color: T.textSecondary, marginTop: 6, flexWrap: 'wrap' }}>
        <span>
          <svg width="22" height="8" aria-hidden="true" style={{ marginRight: 4 }}>
            <line x1="0" y1="4" x2="22" y2="4" strokeWidth={2} strokeDasharray="6 4" style={{ stroke: 'var(--neutralTertiary, #8A8886)' }} />
          </svg>
          {strings.View_Dashboard_BurndownIdeal}
        </span>
        <span>
          <svg width="22" height="8" aria-hidden="true" style={{ marginRight: 4 }}>
            <line x1="0" y1="4" x2="22" y2="4" strokeWidth={2.5} stroke={color} />
          </svg>
          {strings.View_Dashboard_BurndownActual}
        </span>
        <span style={{ marginLeft: 'auto' }}>{summary}</span>
      </div>
    </div>
  );
};

// ── Main component ─────────────────────────────────────────────────────────
export const DashboardView: React.FC<IDashboardViewProps> = ({
  project, tasks, onEditTask, onAddTask, calendar,
}) => {
  const today = todayLocalMidnight();

  // ── Aggregate stats ──────────────────────────────────────────────────────
  const total = tasks.length;
  const byStatus = React.useMemo(() => {
    const counts: Record<string, number> = {
      'Completed': 0, 'In Progress': 0, 'Not Started': 0, 'On Hold': 0, 'Cancelled': 0,
    };
    tasks.forEach(t => { if (counts[t.status] !== undefined) counts[t.status]++; });
    return counts;
  }, [tasks]);

  // Progress basis: leaf tasks only (a parent's own % is not meaningful next
  // to its sub-tasks and would double count them) and never Cancelled tasks
  // (they will not be done, so they would drag the percentage down forever).
  // Each counted task weighs the same — an unweighted mean of % complete.
  const progressTasks = React.useMemo(() => {
    const parentIds = new Set<number>();
    tasks.forEach(t => { if (t.parentTaskId) parentIds.add(t.parentTaskId); });
    return tasks.filter(t => t.status !== 'Cancelled' && !parentIds.has(t.id));
  }, [tasks]);

  const overallPct = progressTasks.length > 0
    ? Math.round(progressTasks.reduce((s, t) => s + t.percentComplete, 0) / progressTasks.length)
    : 0;

  // ── Phase progress ───────────────────────────────────────────────────────
  const phases = React.useMemo(() => {
    const phaseMap = new Map<string, ITask[]>();
    progressTasks.forEach(t => {
      if (!t.phase) return;
      if (!phaseMap.has(t.phase)) phaseMap.set(t.phase, []);
      phaseMap.get(t.phase)!.push(t);
    });
    return Array.from(phaseMap.entries()).map(([name, pts]) => ({
      name,
      color: phaseColor(name),
      total: pts.length,
      completed: pts.filter(t => t.status === 'Completed').length,
      pct: Math.round(pts.reduce((s, t) => s + t.percentComplete, 0) / pts.length),
    }));
  }, [progressTasks]);

  // ── Time-based groups ────────────────────────────────────────────────────
  const { recentlyModified, recentlyCompleted, upcoming, overdue } = React.useMemo(() => {
    const in14 = addDays(today, 14);
    const weekAgo = addDays(today, -7);
    const byModifiedDesc = (a: ITask, b: ITask): number =>
      (parseTimestamp(b.modified)?.getTime() ?? 0) - (parseTimestamp(a.modified)?.getTime() ?? 0);
    const dueTime = (t: ITask): number => parseDateOnly(t.dueDate)?.getTime() ?? 0;

    const recentlyModified = tasks.filter(t => {
      const mod = parseTimestamp(t.modified);
      return mod !== null && mod >= weekAgo && t.status !== 'Completed';
    }).sort(byModifiedDesc).slice(0, 6);

    const recentlyCompleted = tasks.filter(t => {
      const mod = parseTimestamp(t.modified);
      return mod !== null && mod >= weekAgo && t.status === 'Completed';
    }).sort(byModifiedDesc).slice(0, 4);

    const upcoming = tasks.filter(t => {
      const due = parseDateOnly(t.dueDate);
      return due !== null && due >= today && due <= in14 && t.status !== 'Completed' && t.status !== 'Cancelled';
    }).sort((a, b) => dueTime(a) - dueTime(b)).slice(0, 6);

    // Most overdue first: the oldest due date has been late the longest.
    const overdue = tasks.filter(t => {
      const due = parseDateOnly(t.dueDate);
      return due !== null && due < today && t.status !== 'Completed' && t.status !== 'Cancelled';
    }).sort((a, b) => dueTime(a) - dueTime(b)).slice(0, 6);

    return { recentlyModified, recentlyCompleted, upcoming, overdue };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tasks, today.getTime()]);

  // ── Priority breakdown ───────────────────────────────────────────────────
  const byPriority = React.useMemo(() => {
    const counts: Record<string, number> = { Critical: 0, High: 0, Medium: 0, Low: 0 };
    tasks.forEach(t => { if (counts[t.priority] !== undefined) counts[t.priority]++; });
    return counts;
  }, [tasks]);

  const priorityItems = [
    { label: 'Critical', color: PRIORITY_COLORS['Critical'] },
    { label: 'High',     color: PRIORITY_COLORS['High']     },
    { label: 'Medium',   color: PRIORITY_COLORS['Medium']   },
    { label: 'Low',      color: PRIORITY_COLORS['Low']      },
  ];

  // ── Burndown ─────────────────────────────────────────────────────────────
  const burndown = React.useMemo(
    () => buildBurndown(progressTasks, project, today),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [progressTasks, project.startDate, project.dueDate, today.getTime()],
  );

  // ── Workload per assignee (open leaf tasks) ──────────────────────────────
  const workload = React.useMemo(() => {
    const map = new Map<string, { open: number; overdue: number }>();
    progressTasks.forEach(t => {
      if (t.status === 'Completed') return;
      const key = t.assignedTo || '';
      const entry = map.get(key) || { open: 0, overdue: 0 };
      entry.open++;
      const due = parseDateOnly(t.dueDate);
      if (due && due < today) entry.overdue++;
      map.set(key, entry);
    });
    return Array.from(map.entries())
      .map(([name, v]) => ({ name, ...v }))
      .sort((a, b) => b.open - a.open || a.name.localeCompare(b.name))
      .slice(0, 8);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [progressTasks, today.getTime()]);
  const maxOpen = workload.reduce((m, w) => Math.max(m, w.open), 0);

  // ── Critical path ────────────────────────────────────────────────────────
  // Shown only when it is an actual chain (2+ tasks); with no dependencies
  // there is nothing meaningful to call "critical".
  const criticalTasks = React.useMemo(() => {
    const ids = computeCriticalPath(tasks, calendar || DEFAULT_WORKING_CALENDAR);
    return tasks
      .filter(t => ids.has(t.id))
      .sort((a, b) => (parseDateOnly(a.startDate)?.getTime() ?? 0) - (parseDateOnly(b.startDate)?.getTime() ?? 0));
  }, [tasks, calendar]);

  const panel: React.CSSProperties = { background: T.surface, borderRadius: 8, border: `1px solid ${T.border}`, padding: '16px 20px' };

  return (
    <div style={{ padding: '20px 24px', overflowY: 'auto', background: T.surfaceAlt, minHeight: '100%' }}>
      <style>{SHARED_VIEW_CSS}</style>

      {/* ── Project header ─────────────────────────────────────────────── */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 20 }}>
        <div className="sgFill" style={{ width: 14, height: 14, borderRadius: '50%', background: project.color, flexShrink: 0 }} />
        <div>
          <div style={{ fontSize: 20, fontWeight: 700, color: T.text, lineHeight: 1.2 }}>{project.title}</div>
          {project.description && (
            <div style={{ fontSize: 12, color: T.textSecondary, marginTop: 2 }}>{project.description}</div>
          )}
        </div>
        <div style={{ marginLeft: 'auto', display: 'flex', gap: 8 }}>
          {project.startDate && project.dueDate && (
            <div style={{ fontSize: 12, color: T.textSecondary, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: '4px 10px' }}>
              {fmtDate(project.startDate)} → {fmtDate(project.dueDate)}
            </div>
          )}
          <div style={{
            fontSize: 12, fontWeight: 600, padding: '4px 10px', borderRadius: 4,
            background: PROJECT_STATUS_LIGHT_COLORS[project.status] || T.surfaceMuted,
            color: PROJECT_STATUS_COLORS[project.status] || T.textSecondary,
          }}>
            {getProjectStatusLabel(project.status)}
          </div>
        </div>
      </div>

      {/* ── Stat cards ─────────────────────────────────────────────────── */}
      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginBottom: 20 }}>
        <StatCard label={strings.Dashboard_TotalTasks}  value={total}                   total={0}     color="#323130" bg="#F3F2F1" />
        <StatCard label={strings.Dashboard_Completed}    value={byStatus['Completed']}   total={total} color={STATUS_COLORS['Completed']}   bg={STATUS_LIGHT_COLORS['Completed']}   />
        <StatCard label={strings.Dashboard_InProgress}  value={byStatus['In Progress']} total={total} color={STATUS_COLORS['In Progress']} bg={STATUS_LIGHT_COLORS['In Progress']} />
        <StatCard label={strings.Dashboard_OnHold}      value={byStatus['On Hold']}     total={total} color={STATUS_COLORS['On Hold']}     bg={STATUS_LIGHT_COLORS['On Hold']}     />
        <StatCard label={strings.Dashboard_NotStarted}  value={byStatus['Not Started']} total={total} color={STATUS_COLORS['Not Started']} bg={STATUS_LIGHT_COLORS['Not Started']} />
        {byStatus['Cancelled'] > 0 && (
          <StatCard label={strings.Dashboard_Cancelled}  value={byStatus['Cancelled']}   total={total} color={STATUS_COLORS['Cancelled']}   bg={STATUS_LIGHT_COLORS['Cancelled']}   />
        )}
      </div>

      {/* ── Overall progress ───────────────────────────────────────────── */}
      <div className="sgCard" style={{ ...panel, marginBottom: 16 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
          <span style={{ fontSize: 13, fontWeight: 700, color: T.text }}>{strings.Dashboard_OverallProgress}</span>
          <span style={{ fontSize: 16, fontWeight: 700, color: project.color }}>{overallPct}%</span>
        </div>
        <div
          role="progressbar"
          aria-label={strings.Dashboard_OverallProgress}
          aria-valuemin={0}
          aria-valuemax={100}
          aria-valuenow={overallPct}
          style={{ height: 10, background: T.border, borderRadius: 5, overflow: 'hidden' }}
        >
          <div className="sgFill" style={{ height: '100%', width: `${overallPct}%`, background: project.color, borderRadius: 5, transition: 'width 0.4s ease' }} />
        </div>
        {project.projectManager && (
          <div style={{ fontSize: 11, color: T.textTertiary, marginTop: 8 }}>{formatString(strings.Dashboard_ProjectManagerLabel, { name: project.projectManager })}</div>
        )}
      </div>

      {/* ── Phase progress + Status / Priority breakdown ────────────────── */}
      <div style={{ display: 'grid', gridTemplateColumns: phases.length ? '1fr 1fr' : '1fr', gap: 16, marginBottom: 16 }}>

        {/* Phase progress */}
        {phases.length > 0 && (
          <div className="sgCard" style={panel}>
            <SectionHeader title={strings.Dashboard_PhaseProgressHeader} />
            {phases.map(ph => (
              <div key={ph.name} style={{ marginBottom: 12 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 4 }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <div className="sgFill" style={{ width: 8, height: 8, borderRadius: '50%', background: ph.color }} />
                    <span style={{ fontSize: 12, color: T.text, fontWeight: 500 }}>{ph.name}</span>
                  </div>
                  <span style={{ fontSize: 11, color: T.textSecondary }}>
                    {formatString(strings.View_Dashboard_PhaseDoneSummary, { done: ph.completed, total: ph.total, percent: ph.pct })}
                  </span>
                </div>
                <div
                  role="progressbar"
                  aria-label={ph.name}
                  aria-valuemin={0}
                  aria-valuemax={100}
                  aria-valuenow={ph.pct}
                  style={{ height: 6, background: T.border, borderRadius: 3, overflow: 'hidden' }}
                >
                  <div className="sgFill" style={{ height: '100%', width: `${ph.pct}%`, background: ph.color, borderRadius: 3 }} />
                </div>
              </div>
            ))}
          </div>
        )}

        {/* Status + Priority breakdown */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
          <div className="sgCard" style={{ ...panel, flex: 1 }}>
            <SectionHeader title={strings.Dashboard_StatusBreakdownHeader} />
            {(['Completed','In Progress','Not Started','On Hold','Cancelled'] as TaskStatus[]).map(s => {
              const cnt = byStatus[s] || 0;
              const pct = total > 0 ? Math.round(cnt / total * 100) : 0;
              return (
                <div key={s} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 7 }}>
                  <div className="sgFill" style={{ width: 8, height: 8, borderRadius: '50%', background: STATUS_COLORS[s], flexShrink: 0 }} />
                  <span style={{ fontSize: 12, color: T.text, flex: 1 }}>{getStatusLabel(s)}</span>
                  <span style={{ fontSize: 12, fontWeight: 600, color: STATUS_COLORS[s], width: 24, textAlign: 'right' }}>{cnt}</span>
                  <div style={{ width: 80, height: 4, background: T.border, borderRadius: 2, overflow: 'hidden' }}>
                    <div className="sgFill" style={{ height: '100%', width: `${pct}%`, background: STATUS_COLORS[s], borderRadius: 2 }} />
                  </div>
                  <span style={{ fontSize: 10, color: T.textTertiary, width: 30, textAlign: 'right' }}>{pct}%</span>
                </div>
              );
            })}
          </div>

          <div className="sgCard" style={panel}>
            <SectionHeader title={strings.Dashboard_PriorityBreakdownHeader} />
            {priorityItems.map(p => {
              const cnt = byPriority[p.label] || 0;
              const pct = total > 0 ? Math.round(cnt / total * 100) : 0;
              return (
                <div key={p.label} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 7 }}>
                  <div className="sgFill" style={{ width: 8, height: 8, borderRadius: '50%', background: p.color, flexShrink: 0 }} />
                  <span style={{ fontSize: 12, color: T.text, flex: 1 }}>{getPriorityLabel(p.label as TaskPriority)}</span>
                  <span style={{ fontSize: 12, fontWeight: 600, color: p.color, width: 24, textAlign: 'right' }}>{cnt}</span>
                  <div style={{ width: 80, height: 4, background: T.border, borderRadius: 2, overflow: 'hidden' }}>
                    <div className="sgFill" style={{ height: '100%', width: `${pct}%`, background: p.color, borderRadius: 2 }} />
                  </div>
                  <span style={{ fontSize: 10, color: T.textTertiary, width: 30, textAlign: 'right' }}>{pct}%</span>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* ── Burndown ────────────────────────────────────────────────────── */}
      {total > 0 && (
        <div className="sgCard" style={{ ...panel, marginBottom: 16 }}>
          <SectionHeader title={strings.View_Dashboard_BurndownHeader} />
          {burndown ? (
            <BurndownChart data={burndown} color={project.color || '#0078D4'} />
          ) : (
            <div style={{ fontSize: 12, color: T.textSecondary }}>{strings.View_Dashboard_BurndownNoData}</div>
          )}
        </div>
      )}

      {/* ── Workload + Critical path ────────────────────────────────────── */}
      {(workload.length > 0 || criticalTasks.length > 1) && (
        <div style={{ display: 'grid', gridTemplateColumns: workload.length > 0 && criticalTasks.length > 1 ? '1fr 1fr' : '1fr', gap: 16, marginBottom: 16 }}>

          {workload.length > 0 && (
            <div className="sgCard" style={panel}>
              <SectionHeader title={strings.View_Dashboard_WorkloadHeader} />
              {workload.map(w => {
                const label = w.name || strings.ListView_Unassigned;
                return (
                  <div key={w.name || '__unassigned'} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                    <div
                      aria-hidden="true"
                      style={{
                        width: 22, height: 22, borderRadius: '50%', flexShrink: 0,
                        background: w.name ? stringToColor(w.name) : '#8B929A', color: '#fff',
                        fontSize: 10, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center',
                      }}
                    >
                      {w.name ? initials(w.name) : '?'}
                    </div>
                    <span style={{ fontSize: 12, color: T.text, width: 110, flexShrink: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={label}>
                      {label}
                    </span>
                    <div
                      role="img"
                      aria-label={formatString(strings.View_Dashboard_WorkloadOpen, { count: w.open })}
                      style={{ flex: 1, height: 8, background: T.border, borderRadius: 4, overflow: 'hidden' }}
                    >
                      <div className="sgFill" style={{ height: '100%', width: `${maxOpen > 0 ? Math.round(w.open / maxOpen * 100) : 0}%`, background: project.color, borderRadius: 4 }} />
                    </div>
                    <span style={{ fontSize: 11, color: T.textSecondary, width: 64, textAlign: 'right', flexShrink: 0 }}>
                      {formatString(strings.View_Dashboard_WorkloadOpen, { count: w.open })}
                    </span>
                    {w.overdue > 0 && (
                      <span style={{ fontSize: 11, color: '#D13438', fontWeight: 600, flexShrink: 0 }}>
                        {formatString(strings.View_Dashboard_WorkloadOverdue, { count: w.overdue })}
                      </span>
                    )}
                  </div>
                );
              })}
            </div>
          )}

          {criticalTasks.length > 1 && (
            <div className="sgCard" style={panel}>
              <SectionHeader title={formatString(strings.View_Dashboard_CriticalPathHeader, { count: criticalTasks.length })} />
              {criticalTasks.map(t => (
                <TaskRow key={t.id} task={t} onClick={() => onEditTask(t)} showDue />
              ))}
            </div>
          )}

        </div>
      )}

      {/* ── Overdue + Upcoming ──────────────────────────────────────────── */}
      {(overdue.length > 0 || upcoming.length > 0) && (
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 16 }}>

          {overdue.length > 0 && (
            <div className="sgCard" style={panel}>
              <SectionHeader title={formatString(strings.Dashboard_OverdueHeader, { count: overdue.length })} />
              {overdue.map(t => (
                <TaskRow key={t.id} task={t} onClick={() => onEditTask(t)} showDue isOverdue />
              ))}
            </div>
          )}

          {upcoming.length > 0 && (
            <div className="sgCard" style={panel}>
              <SectionHeader title={formatString(strings.Dashboard_DueNext14DaysHeader, { count: upcoming.length })} />
              {upcoming.map(t => {
                const due = parseDateOnly(t.dueDate)!;
                const daysLeft = differenceInCalendarDays(due, today);
                return (
                  <div
                    key={t.id}
                    className="sgFocusable"
                    role="button"
                    tabIndex={0}
                    aria-label={formatString(strings.View_Dashboard_TaskRowAriaLabel, { taskTitle: t.title, status: getStatusLabel(t.status as TaskStatus) })}
                    style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4, cursor: 'pointer', padding: '6px 8px', borderRadius: 5 }}
                    onClick={() => onEditTask(t)}
                    onKeyDown={onActivate(() => onEditTask(t))}
                    onMouseEnter={e => { (e.currentTarget as HTMLElement).style.background = 'var(--neutralLighter, #F3F2F1)'; }}
                    onMouseLeave={e => { (e.currentTarget as HTMLElement).style.background = 'transparent'; }}
                  >
                    <div className="sgFill" style={{ width: 8, height: 8, borderRadius: '50%', background: STATUS_COLORS[t.status as TaskStatus] || '#8B929A', flexShrink: 0 }} />
                    <span style={{ fontSize: 13, color: T.text, flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {t.isMilestone ? '◆ ' : ''}{t.title}
                    </span>
                    <span style={{ fontSize: 11, color: daysLeft <= 3 ? '#CA5010' : T.textSecondary, fontWeight: daysLeft <= 3 ? 600 : 400, flexShrink: 0 }}>
                      {daysLeft === 0 ? strings.Dashboard_Today : daysLeft === 1 ? strings.Dashboard_Tomorrow : formatString(strings.Dashboard_DaysLeft, { days: daysLeft })}
                    </span>
                  </div>
                );
              })}
            </div>
          )}

        </div>
      )}

      {/* ── Recent activity ─────────────────────────────────────────────── */}
      {(recentlyCompleted.length > 0 || recentlyModified.length > 0) && (
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 16 }}>

          {recentlyCompleted.length > 0 && (
            <div className="sgCard" style={panel}>
              <SectionHeader title={strings.Dashboard_CompletedThisWeekHeader} />
              {recentlyCompleted.map(t => (
                <TaskRow key={t.id} task={t} onClick={() => onEditTask(t)} />
              ))}
            </div>
          )}

          {recentlyModified.length > 0 && (
            <div className="sgCard" style={panel}>
              <SectionHeader title={strings.Dashboard_UpdatedThisWeekHeader} />
              {recentlyModified.map(t => (
                <TaskRow key={t.id} task={t} onClick={() => onEditTask(t)} />
              ))}
            </div>
          )}

        </div>
      )}

      {/* ── Empty state ─────────────────────────────────────────────────── */}
      {total === 0 && (
        <div style={{ textAlign: 'center', padding: '60px 0', color: T.textSecondary }}>
          <div style={{ fontSize: 32, marginBottom: 12 }}>📋</div>
          <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 8, color: T.text }}>{strings.Dashboard_EmptyTitle}</div>
          <div style={{ fontSize: 13, marginBottom: 20 }}>{strings.Dashboard_EmptySubtitle}</div>
          <button
            onClick={onAddTask}
            style={{
              padding: '8px 20px', background: project.color, color: '#fff',
              border: 'none', borderRadius: 4, fontSize: 13, fontWeight: 600, cursor: 'pointer',
            }}
          >
            {strings.Dashboard_AddTaskButton}
          </button>
        </div>
      )}

    </div>
  );
};

export default DashboardView;
