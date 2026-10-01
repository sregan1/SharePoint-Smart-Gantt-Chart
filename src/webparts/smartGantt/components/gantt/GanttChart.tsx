import * as React from 'react';
import {
  addDays, addMonths, addWeeks, differenceInCalendarDays,
  endOfMonth, format, max, min,
  startOfMonth, startOfWeek, getISOWeek,
} from 'date-fns';
import {
  IProject, ITask, ZoomLevel, IGanttDisplaySettings, DEFAULT_GANTT_SETTINGS,
  HEADER_THEME_COLORS, IWorkingCalendar, DEFAULT_WORKING_CALENDAR,
} from '../../models';
import { parseDateOnly, formatDateOnly, dateToDateOnlyString, todayLocalMidnight } from '../../utils/dateUtils';
import { computeCriticalPath, isWorkingDay } from '../../utils/scheduleUtils';
import { formatDependencyLink } from '../../utils/dependencyUtils';
import * as strings from 'SmartGanttWebPartStrings';
import { formatString } from '../localeUtils';
import { GanttBar, IBarHandlers, BarDragMode } from './GanttBar';
import { GanttTooltip, ITooltipApi } from './GanttTooltip';
import {
  HEADER_HEIGHT, BAR_HEIGHT, MILESTONE_SIZE, MIN_BAR_WIDTH, BUFFER_DAYS, ROW_OVERSCAN,
  DAY_WIDTH, CRITICAL_COLOR, zoomForDayWidth, getTaskColor, colorId, taskDuration,
  snapToWorkingDay, shiftWorkingDays, linkOf, isLinkViolated, buildVisibleRows,
  IVisibleRow, ITaskUpdateOptions, loadCollapsedPhases, saveCollapsedPhases,
} from './ganttUtils';
import styles from './GanttChart.module.scss';

interface IGanttChartProps {
  tasks: ITask[];
  project: IProject;
  zoomLevel: ZoomLevel;
  settings: IGanttDisplaySettings;
  scrollToToday: boolean;
  /** Working days + holidays: drives snapping, shading and keyboard moves. */
  calendar?: IWorkingCalendar;
  /** Custom day width (fit-to-project / Ctrl+wheel). Null/absent = use zoomLevel. */
  dayWidthOverride?: number | null;
  onDayWidthChange?: (dayWidth: number | null) => void;
  /** Increment to ask the chart to fit the whole project into view. */
  fitSignal?: number;
  onEditTask: (task: ITask) => void;
  onDeleteTask: (id: number) => void;
  onTaskUpdate: (id: number, updates: Partial<ITask>, opts?: ITaskUpdateOptions) => void;
  onAddTask: () => void;
  onImport?: () => void;
  /** Drag across an empty row to create a task with these ISO dates. */
  onCreateTaskAt?: (startDate: string, dueDate: string) => void;
  /** Drag from a bar's connector to another bar: predecessor → successor. */
  onAddDependency?: (predId: number, succId: number) => void;
}

type DragMode = BarDragMode | 'link' | 'create';

interface IDragState {
  taskId: number;
  mode: DragMode;
  pointerId: number;
  startClientX: number;
  startClientY: number;
  startScrollLeft: number;
  origStartDate: Date;
  origEndDate: Date;
  /** link: SVG coordinates of the connector; create: row index. */
  fromX?: number;
  fromY?: number;
  row?: number;
}

interface IViewport { left: number; top: number; width: number; height: number }

interface IGhost { row: number; start: Date; end: Date }

// Everything the stable pointer/keyboard handlers need, refreshed each render.
interface ILatest {
  tasks: ITask[];
  taskById: Map<number, ITask>;
  rows: IVisibleRow[];
  dayWidth: number;
  rowH: number;
  rangeStart: Date;
  cal: IWorkingCalendar;
  today: Date;
  dependenciesOnHover: boolean;
  dragging: boolean;
  onTaskUpdate: IGanttChartProps['onTaskUpdate'];
  onEditTask: IGanttChartProps['onEditTask'];
  onCreateTaskAt?: IGanttChartProps['onCreateTaskAt'];
  onAddDependency?: IGanttChartProps['onAddDependency'];
}

let ganttInstanceCounter = 0;

// ─── Left panel row ──────────────────────────────────────────────────────────

interface ILeftRowProps {
  row: IVisibleRow;
  rowH: number;
  color: string;
  violated: boolean;
  collapsed: boolean;
  onEditTask: (task: ITask) => void;
  onDeleteTask: (id: number) => void;
  onTogglePhase: (phase: string) => void;
}

const LeftRowComponent: React.FC<ILeftRowProps> = ({
  row, rowH, color, violated, collapsed, onEditTask, onDeleteTask, onTogglePhase,
}) => {
  if (row.type === 'phase') {
    return (
      <div className={`${styles.taskRow} ${styles.phaseRow}`} style={{ height: rowH }}>
        <button
          className={styles.taskExpandBtn}
          aria-label={formatString(strings.Gantt_ExpandPhaseAriaLabel, {
            action: collapsed ? strings.Gantt_ExpandAction : strings.Gantt_CollapseAction,
            phase: row.phase!,
          })}
          aria-expanded={!collapsed}
          onClick={() => onTogglePhase(row.phase!)}
        >
          {collapsed ? '▶' : '▼'}
        </button>
        <span className={styles.phaseLabel}>&ensp;{row.phase}</span>
      </div>
    );
  }

  const task = row.task!;
  const dur = taskDuration(task);
  return (
    <div className={styles.taskRow} style={{ height: rowH }}>
      {row.depth ? <div className={styles.taskIndent} style={{ width: 16 * row.depth }} /> : null}
      <div className={styles.taskStatusDot} style={{ background: color }} />
      {task.isMilestone && <span className={styles.milestoneIcon}>◆</span>}
      <span className={styles.taskName} title={task.title}>{task.title}</span>
      {row.cyclic && (
        <span
          className={styles.rowWarning}
          title={strings.Gantt_ParentCycleTooltip}
          role="img"
          aria-label={strings.Gantt_ParentCycleTooltip}
        >
          ⚠
        </span>
      )}
      {violated && (
        <span
          className={styles.rowWarning}
          title={strings.Gantt_DependencyViolationTooltip}
          role="img"
          aria-label={strings.Gantt_DependencyViolationTooltip}
        >
          ⚠
        </span>
      )}
      <span className={styles.taskDuration}>{dur > 0 ? `${dur}d` : '—'}</span>
      <div className={styles.taskRowActions}>
        <button
          className={styles.taskActionBtn}
          onClick={e => { e.stopPropagation(); onEditTask(task); }}
          title={strings.Gantt_EditTitle}
          aria-label={formatString(strings.Gantt_EditTaskAriaLabel, { taskTitle: task.title })}
        >
          ✏
        </button>
        <button
          className={`${styles.taskActionBtn} ${styles.deleteBtn}`}
          onClick={e => { e.stopPropagation(); onDeleteTask(task.id); }}
          title={strings.Gantt_DeleteTitle}
          aria-label={formatString(strings.Gantt_DeleteTaskAriaLabel, { taskTitle: task.title })}
        >
          ✕
        </button>
      </div>
    </div>
  );
};

const LeftRow = React.memo(LeftRowComponent);

// ─── Chart ───────────────────────────────────────────────────────────────────

export const GanttChart: React.FC<IGanttChartProps> = ({
  tasks,
  project,
  zoomLevel,
  settings = DEFAULT_GANTT_SETTINGS,
  scrollToToday,
  calendar,
  dayWidthOverride,
  onDayWidthChange,
  fitSignal,
  onEditTask,
  onDeleteTask,
  onTaskUpdate,
  onAddTask,
  onImport,
  onCreateTaskAt,
  onAddDependency,
}) => {
  const bodyScrollRef = React.useRef<HTMLDivElement>(null);
  const headerScrollRef = React.useRef<HTMLDivElement>(null);
  const leftBodyRef = React.useRef<HTMLDivElement>(null);
  const svgRef = React.useRef<SVGSVGElement>(null);
  const tooltipApiRef = React.useRef<ITooltipApi | null>(null);
  const dragLabelRef = React.useRef<HTMLDivElement>(null);
  const linkLineRef = React.useRef<SVGLineElement>(null);
  const latest = React.useRef<ILatest>({} as ILatest);
  // The ghost is state (it renders), but the pointerup closure needs the latest value.
  const lastGhostRef = React.useRef<IGhost | null>(null);

  // Unique per-instance prefix so SVG defs (gradients, markers) don't collide
  // when two Smart Gantt web parts render on the same page. Lazily
  // initialized so the module counter increments once per mount, not once
  // per render (React.useRef(expr) evaluates expr on every render even
  // though only the first value is kept).
  const uidRef = React.useRef<string>();
  if (!uidRef.current) uidRef.current = `sg${++ganttInstanceCounter}`;
  const uid = uidRef.current;
  const hintId = `${uid}-hint`;

  const cal = calendar || DEFAULT_WORKING_CALENDAR;

  const [dragState, setDragState] = React.useState<IDragState | null>(null);
  const [dragOffsets, setDragOffsets] = React.useState<Map<number, { start: Date; end: Date }>>(new Map());
  // Live mirror of dragOffsets so the pointer listeners don't need
  // re-registering on every state change.
  const dragOffsetsRef = React.useRef(dragOffsets);
  // Set after a real drag so the click that the browser fires on pointerup
  // doesn't also open the edit panel.
  const suppressClickRef = React.useRef(false);
  const [hoveredId, setHoveredId] = React.useState<number | null>(null);
  const [linkTargetId, setLinkTargetId] = React.useState<number | null>(null);
  const [ghost, setGhost] = React.useState<IGhost | null>(null);
  const [announce, setAnnounce] = React.useState('');
  const [collapsedPhases, setCollapsedPhases] = React.useState<Set<string>>(
    () => loadCollapsedPhases(project.listName)
  );
  const [view, setView] = React.useState<IViewport>({ left: 0, top: 0, width: 1200, height: 700 });
  // Tracks whether the previous render had scrollToToday=true, so the
  // false-reset 100ms later doesn't override the Today scroll with the
  // smart-initial-position logic.
  const prevScrollToToday = React.useRef(false);
  const scrollLeftRef = React.useRef(0);
  const prevGeomRef = React.useRef<{ start: Date; dw: number } | null>(null);
  const zoomAnchorRef = React.useRef<{ offsetX: number; dayFloat: number } | null>(null);
  const pendingFitScrollRef = React.useRef<Date | null>(null);
  const lastFitSignalRef = React.useRef(fitSignal || 0);
  const viewRafRef = React.useRef(0);

  const dayWidth = dayWidthOverride || DAY_WIDTH[zoomLevel];
  // Header layout follows the effective zoom so fit / Ctrl+wheel widths still
  // get a sensible header (day numbers vs week bands vs month bands).
  const effZoom: ZoomLevel = dayWidthOverride ? zoomForDayWidth(dayWidth) : zoomLevel;
  const today = todayLocalMidnight();
  const todayMs = today.getTime();
  const ROW_H = settings.rowHeight;
  const theme = HEADER_THEME_COLORS[settings.headerTheme];

  // For project-relative week numbers: find earliest task start
  const projectWeekStart = React.useMemo(() => {
    const earliest = tasks.reduce<Date | null>((acc, t) => {
      const s = parseDateOnly(t.startDate);
      return s && (!acc || s < acc) ? s : acc;
    }, null);
    return startOfWeek(earliest || new Date(todayMs), { weekStartsOn: 1 });
  }, [tasks, todayMs]);

  const criticalIds = React.useMemo(
    () => (settings.showCriticalPath || settings.showCriticalPathOnly ? computeCriticalPath(tasks, cal) : new Set<number>()),
    [tasks, cal, settings.showCriticalPath, settings.showCriticalPathOnly]
  );

  // Successors whose committed dates break at least one of their links.
  const violationIds = React.useMemo(() => {
    const byId = new Map(tasks.map(t => [t.id, t]));
    const set = new Set<number>();
    tasks.forEach(t => {
      const ts = parseDateOnly(t.startDate);
      const te = parseDateOnly(t.dueDate);
      if (!ts || !te) return;
      t.dependencies.forEach(pid => {
        const p = byId.get(pid);
        if (!p) return;
        const ps = parseDateOnly(p.startDate);
        const pe = parseDateOnly(p.dueDate);
        if (!ps || !pe) return;
        if (isLinkViolated({ start: ps, end: pe }, { start: ts, end: te }, linkOf(t, pid), cal)) set.add(t.id);
      });
    });
    return set;
  }, [tasks, cal]);

  // One gradient per distinct bar color rather than one per task — with
  // hundreds of tasks sharing the same status/phase/priority palette this
  // keeps the <defs> count near the palette size instead of the task count.
  const distinctBarColors = React.useMemo(
    () => Array.from(new Set(tasks.map(t => getTaskColor(t, settings)))),
    [tasks, settings]
  );

  const getWeekLabel = (weekDate: Date): string => {
    if (settings.weekLabel === 'dates') return format(weekDate, 'MMM d');
    if (settings.weekLabel === 'project') {
      const diff = differenceInCalendarDays(weekDate, projectWeekStart);
      // The BUFFER_DAYS lead-in before the earliest task falls before week 1;
      // clamping it to "W1" made several consecutive header bands share that
      // label, indistinguishable from the real week 1.
      if (diff < 0) return '';
      const wn = Math.floor(diff / 7) + 1;
      return `W${wn}`;
    }
    return `W${getISOWeek(weekDate)}`;
  };

  // Compute visible date range
  const { rangeStart, rangeEnd } = React.useMemo(() => {
    const t0 = new Date(todayMs);
    const dates: Date[] = [addDays(t0, -BUFFER_DAYS)];
    tasks.forEach(t => {
      const s = parseDateOnly(t.startDate);
      const e = parseDateOnly(t.dueDate);
      if (s) dates.push(addDays(s, -BUFFER_DAYS));
      if (e) dates.push(addDays(e, BUFFER_DAYS));
    });
    const earliest = startOfMonth(dates.reduce((a, b) => (a < b ? a : b), t0));
    const latestD = endOfMonth(dates.reduce((a, b) => (a > b ? a : b), t0));
    return { rangeStart: earliest, rangeEnd: latestD };
  }, [tasks, todayMs]);

  const totalDays = differenceInCalendarDays(rangeEnd, rangeStart) + 1;
  const svgWidth = totalDays * dayWidth;
  // Rebuilding this per render is otherwise the most expensive step in the
  // component — it runs on every pointer move during a drag, neither of which
  // changes tasks/collapsedPhases.
  const visibleTasks = React.useMemo(
    () => buildVisibleRows(tasks, collapsedPhases),
    [tasks, collapsedPhases]
  );
  const svgBodyHeight = visibleTasks.length * ROW_H;
  const bodyHeight = Math.max(svgBodyHeight + 60, 400);

  const taskById = React.useMemo(() => new Map(tasks.map(t => [t.id, t])), [tasks]);
  const rowIndexById = React.useMemo(() => {
    const map = new Map<number, number>();
    // Row positions must come from the *visible* rows (phase headers shift
    // and reorder everything), not from the raw tasks array.
    visibleTasks.forEach((row, i) => {
      if (row.type === 'task' && row.task) map.set(row.task.id, i);
    });
    return map;
  }, [visibleTasks]);

  const phaseNames = React.useMemo(
    () => Array.from(new Set(tasks.map(t => t.phase).filter(Boolean))),
    [tasks]
  );

  // Convert date ↔ x
  const dateToX = (d: Date): number => differenceInCalendarDays(d, rangeStart) * dayWidth;

  // Row window (virtualization): only rows near the viewport get DOM nodes.
  const rowStart = Math.max(0, Math.floor(view.top / ROW_H) - ROW_OVERSCAN);
  const rowEnd = Math.min(visibleTasks.length, Math.ceil((view.top + view.height) / ROW_H) + ROW_OVERSCAN + 1);
  // Horizontal window for header ticks (quantized `left` + generous slack).
  const visX0 = view.left - 600;
  const visX1 = view.left + view.width + 800;

  // Refresh the mirror the stable handlers read from.
  latest.current = {
    tasks, taskById, rows: visibleTasks, dayWidth, rowH: ROW_H, rangeStart, cal, today,
    dependenciesOnHover: settings.dependenciesOnHover,
    dragging: !!dragState,
    onTaskUpdate, onEditTask, onCreateTaskAt, onAddDependency,
  };

  // ─── Viewport tracking (scroll → visible window) ────────────────────────

  const updateView = React.useCallback((): void => {
    const el = bodyScrollRef.current;
    const L = latest.current;
    if (!el || !L) return;
    // Quantized so a scroll event only re-renders when the window meaningfully moves.
    const top = Math.floor(el.scrollTop / L.rowH) * L.rowH;
    const left = Math.floor(el.scrollLeft / 200) * 200;
    const width = el.clientWidth;
    const height = el.clientHeight;
    setView(prev => (
      prev.top === top && prev.left === left && prev.width === width && prev.height === height
        ? prev
        : { top, left, width, height }
    ));
  }, []);

  const hasTasks = tasks.length > 0;

  React.useEffect(() => {
    updateView();
    const el = bodyScrollRef.current;
    if (!el || typeof ResizeObserver === 'undefined') return undefined;
    const ro = new ResizeObserver(() => updateView());
    ro.observe(el);
    return () => ro.disconnect();
  }, [hasTasks, ROW_H, dayWidth, visibleTasks.length, updateView]);

  React.useEffect(() => () => {
    if (viewRafRef.current) cancelAnimationFrame(viewRafRef.current);
  }, []);

  // Collapsed phases are remembered per project.
  const firstProjectRef = React.useRef(project.listName);
  React.useEffect(() => {
    if (firstProjectRef.current === project.listName) return;
    firstProjectRef.current = project.listName;
    setCollapsedPhases(loadCollapsedPhases(project.listName));
  }, [project.listName]);

  const updateCollapsed = (next: Set<string>): void => {
    setCollapsedPhases(next);
    saveCollapsedPhases(project.listName, next);
  };

  const togglePhase = React.useCallback((phase: string): void => {
    setCollapsedPhases(prev => {
      const next = new Set(prev);
      if (next.has(phase)) next.delete(phase); else next.add(phase);
      saveCollapsedPhases(project.listName, next);
      return next;
    });
  }, [project.listName]);

  // ─── Scrolling ───────────────────────────────────────────────────────────

  // Keep the same date under the viewport center (or the wheel pointer) when
  // the day width changes, and keep the view steady when the range's start
  // moves (a drag near the edge of the range shifts rangeStart by a month) —
  // otherwise the chart visibly jumps under the user.
  React.useLayoutEffect(() => {
    const prev = prevGeomRef.current;
    prevGeomRef.current = { start: rangeStart, dw: dayWidth };
    const el = bodyScrollRef.current;
    const anchor = zoomAnchorRef.current;
    zoomAnchorRef.current = null;
    if (!prev || !el) return;
    if (prev.dw === dayWidth && prev.start.getTime() === rangeStart.getTime()) return;
    const offsetX = anchor ? anchor.offsetX : el.clientWidth / 2;
    // Read the scroll position remembered from before the DOM changed — the
    // browser may already have clamped el.scrollLeft to the new content width.
    const dayFloat = anchor ? anchor.dayFloat : (scrollLeftRef.current + offsetX) / prev.dw;
    const next = Math.max(0, (dayFloat + differenceInCalendarDays(prev.start, rangeStart)) * dayWidth - offsetX);
    el.scrollLeft = next;
    if (headerScrollRef.current) headerScrollRef.current.scrollLeft = next;
    scrollLeftRef.current = next;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rangeStart.getTime(), dayWidth]);

  // Position the viewport on load / project switch / ◉ Today. Deliberately NOT
  // keyed on the task dates or range: editing or dragging a task shifts those,
  // and re-snapping then would yank the view away from where the user is.
  //   - project or first-data change: scroll to the earliest task so past tasks are visible
  //   - Today button: always snap to today regardless of task dates
  // prevScrollToToday guards against the boolean false-reset 100ms after Today
  // fires overriding the Today scroll with the smart-position logic.
  const scrollKey = `${project.id}:${hasTasks ? 1 : 0}`;
  React.useEffect(() => {
    const el = bodyScrollRef.current;
    if (!el) return;

    const isToday = scrollToToday || prevScrollToToday.current;
    prevScrollToToday.current = scrollToToday;

    let scrollX: number;
    if (isToday) {
      scrollX = Math.max(0, dateToX(today) - 200);
    } else {
      const earliest = tasks.reduce<Date | null>((acc, t) => {
        const s = parseDateOnly(t.startDate);
        return s && (!acc || s < acc) ? s : acc;
      }, null);
      // If the earliest task is more than a week in the past, show it with
      // two weeks of padding so there is timeline context before the first bar.
      // Otherwise fall back to centering on today.
      scrollX = (earliest && earliest < addDays(today, -7))
        ? Math.max(0, dateToX(earliest) - 14 * dayWidth)
        : Math.max(0, dateToX(today) - 200);
    }

    el.scrollLeft = scrollX;
    scrollLeftRef.current = scrollX;
    if (headerScrollRef.current) headerScrollRef.current.scrollLeft = scrollX;
    updateView();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [scrollKey, scrollToToday]);

  // Fit the whole project into the viewport (Toolbar "Fit").
  React.useEffect(() => {
    const sig = fitSignal || 0;
    if (sig === lastFitSignalRef.current) return;
    lastFitSignalRef.current = sig;
    const el = bodyScrollRef.current;
    if (!el || !onDayWidthChange) return;
    let first: Date | null = null;
    let last: Date | null = null;
    for (let i = 0; i < tasks.length; i++) {
      const s = parseDateOnly(tasks[i].startDate) || parseDateOnly(tasks[i].dueDate);
      const e = parseDateOnly(tasks[i].dueDate) || s;
      if (s && (!first || s < first)) first = s;
      if (e && (!last || e > last)) last = e;
    }
    if (!first || !last) return;
    const span = differenceInCalendarDays(last, first) + 1 + 14;
    const target = Math.min(80, Math.max(1.5, (el.clientWidth - 24) / span));
    pendingFitScrollRef.current = first;
    if (Math.abs(target - dayWidth) < 0.01) {
      const x = Math.max(0, dateToX(first) - 7 * dayWidth);
      el.scrollLeft = x;
      if (headerScrollRef.current) headerScrollRef.current.scrollLeft = x;
      scrollLeftRef.current = x;
      pendingFitScrollRef.current = null;
    } else {
      onDayWidthChange(target);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fitSignal]);

  // After a fit-driven width change has rendered, scroll to the project start.
  // Declared after the preserve-position layout effect so it wins.
  React.useLayoutEffect(() => {
    const first = pendingFitScrollRef.current;
    const el = bodyScrollRef.current;
    if (!first || !el) return;
    pendingFitScrollRef.current = null;
    const x = Math.max(0, dateToX(first) - 7 * dayWidth);
    el.scrollLeft = x;
    if (headerScrollRef.current) headerScrollRef.current.scrollLeft = x;
    scrollLeftRef.current = x;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dayWidth]);

  // Ctrl+wheel zooms around the pointer. Needs a non-passive native listener
  // so the browser's own page zoom can be suppressed.
  React.useEffect(() => {
    const el = bodyScrollRef.current;
    if (!el || !onDayWidthChange) return undefined;
    const onWheel = (e: WheelEvent): void => {
      if (!e.ctrlKey && !e.metaKey) return;
      e.preventDefault();
      const L = latest.current;
      if (!L) return;
      const rect = el.getBoundingClientRect();
      const offsetX = e.clientX - rect.left;
      const factor = e.deltaY < 0 ? 1.12 : 1 / 1.12;
      const next = Math.min(80, Math.max(1.5, L.dayWidth * factor));
      if (next === L.dayWidth) return;
      zoomAnchorRef.current = { offsetX, dayFloat: (el.scrollLeft + offsetX) / L.dayWidth };
      onDayWidthChange(next);
    };
    el.addEventListener('wheel', onWheel, { passive: false });
    return () => el.removeEventListener('wheel', onWheel);
  }, [hasTasks, onDayWidthChange]);

  // Sync scroll between panels
  const handleBodyScroll = (): void => {
    const el = bodyScrollRef.current;
    if (!el) return;
    scrollLeftRef.current = el.scrollLeft;
    if (headerScrollRef.current) headerScrollRef.current.scrollLeft = el.scrollLeft;
    if (leftBodyRef.current) leftBodyRef.current.scrollTop = el.scrollTop;
    if (!viewRafRef.current) {
      viewRafRef.current = requestAnimationFrame(() => {
        viewRafRef.current = 0;
        updateView();
      });
    }
  };

  const handleLeftScroll = (): void => {
    if (!leftBodyRef.current || !bodyScrollRef.current) return;
    bodyScrollRef.current.scrollTop = leftBodyRef.current.scrollTop;
  };

  // ─── Interaction handlers (stable identity — memoized bars depend on it) ─

  const handlers = React.useMemo<IBarHandlers>(() => {
    const beginDrag = (
      e: React.PointerEvent,
      base: Omit<IDragState, 'pointerId' | 'startClientX' | 'startClientY' | 'startScrollLeft'>
    ): void => {
      try { (e.currentTarget as Element).setPointerCapture(e.pointerId); } catch { /* capture is best-effort */ }
      suppressClickRef.current = false;
      tooltipApiRef.current?.hide();
      setHoveredId(null);
      setDragState({
        ...base,
        pointerId: e.pointerId,
        startClientX: e.clientX,
        startClientY: e.clientY,
        startScrollLeft: bodyScrollRef.current ? bodyScrollRef.current.scrollLeft : 0,
      });
    };

    return {
      onPointerDown: (e, task, mode) => {
        // Only the primary button / first touch starts a drag — a right-click
        // (context menu) or middle-click must not move the bar.
        if (e.button !== 0 || !e.isPrimary) return;
        e.preventDefault();
        e.stopPropagation();
        const L = latest.current!;
        const s = parseDateOnly(task.startDate) || (task.isMilestone ? parseDateOnly(task.dueDate) : null) || L.today;
        const end = parseDateOnly(task.dueDate) || L.today;
        beginDrag(e, { taskId: task.id, mode, origStartDate: s, origEndDate: end });
      },
      onLinkStart: (e, task, fromX, fromY) => {
        if (e.button !== 0 || !e.isPrimary) return;
        e.preventDefault();
        e.stopPropagation();
        const L = latest.current!;
        beginDrag(e, {
          taskId: task.id, mode: 'link', origStartDate: L.today, origEndDate: L.today, fromX, fromY,
        });
      },
      onActivate: task => {
        if (suppressClickRef.current) {
          suppressClickRef.current = false;
          return;
        }
        latest.current!.onEditTask(task);
      },
      onKeyDown: (e, task) => {
        const L = latest.current!;
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          L.onEditTask(task);
          return;
        }
        if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
        e.preventDefault();
        const dir = e.key === 'ArrowRight' ? 1 : -1;
        const startIso = task.startDate || (task.isMilestone ? task.dueDate : '');
        const dueIso = task.dueDate || startIso;
        if (!startIso || !dueIso) return;
        let newStart = startIso;
        let newDue = dueIso;
        if (e.shiftKey && !task.isMilestone) {
          // Resize the finish by one working day; never before the start.
          newDue = shiftWorkingDays(dueIso, dir, L.cal);
          if (newDue < startIso) return;
        } else {
          // Move the whole task by one working day.
          newStart = shiftWorkingDays(startIso, dir, L.cal);
          newDue = task.isMilestone ? newStart : shiftWorkingDays(dueIso, dir, L.cal);
        }
        if (newStart === startIso && newDue === dueIso) return;
        L.onTaskUpdate(task.id, { startDate: newStart, dueDate: newDue }, { cascade: true });
        setAnnounce(formatString(strings.Gantt_KeyboardMoveAnnouncement, {
          taskTitle: task.title,
          startDate: formatDateOnly(newStart, 'MMM d, yyyy'),
          dueDate: formatDateOnly(newDue, 'MMM d, yyyy'),
        }));
      },
      onHover: (task, x, y) => {
        const L = latest.current!;
        if (L.dragging) return;
        tooltipApiRef.current?.show(task, x, y);
        if (L.dependenciesOnHover) setHoveredId(task.id);
      },
      onLeave: () => {
        tooltipApiRef.current?.hide();
        setHoveredId(null);
      },
    };
  }, []);

  // Drag across an empty part of a task row to create a task there.
  const handleBackgroundPointerDown = (e: React.PointerEvent<SVGRectElement>): void => {
    const L = latest.current!;
    if (e.button !== 0 || !e.isPrimary || !L.onCreateTaskAt || !svgRef.current) return;
    const rect = svgRef.current.getBoundingClientRect();
    const row = Math.floor((e.clientY - rect.top) / L.rowH);
    const r = L.rows[row];
    if (!r || r.type !== 'task') return;
    e.preventDefault();
    try { e.currentTarget.setPointerCapture(e.pointerId); } catch { /* best-effort */ }
    suppressClickRef.current = false;
    const anchor = addDays(L.rangeStart, Math.floor((e.clientX - rect.left) / L.dayWidth));
    setDragState({
      taskId: -1, mode: 'create', pointerId: e.pointerId,
      startClientX: e.clientX, startClientY: e.clientY,
      startScrollLeft: bodyScrollRef.current ? bodyScrollRef.current.scrollLeft : 0,
      origStartDate: anchor, origEndDate: anchor, row,
    });
  };

  // ─── Drag lifecycle ──────────────────────────────────────────────────────
  // Pointer Events (rather than mouse-only) so dragging/resizing bars also
  // works with touch and pen input, which SharePoint pages are commonly
  // viewed on (tablets). The bar captures the pointer on press; the
  // window-level listeners below receive the captured events.

  // Last delta actually applied during the current drag — a day is several
  // pixels wide, so most raw pointer events round to the same result and
  // would otherwise re-render the chart for no visible change.
  const lastKeyRef = React.useRef('');

  React.useEffect(() => {
    if (!dragState) return undefined;
    const ds = dragState;
    lastKeyRef.current = '';
    let lastX = ds.startClientX;
    let lastY = ds.startClientY;
    let raf = 0;

    const setLabel = (text: string, clientX: number, clientY: number): void => {
      const el = dragLabelRef.current;
      if (!el) return;
      el.textContent = text;
      el.style.left = `${clientX + 14}px`;
      el.style.top = `${clientY + 18}px`;
      el.style.display = text ? 'block' : 'none';
    };

    const rangeLabel = (s: Date, e: Date): string => formatString(strings.Gantt_DragLabel, {
      start: format(s, 'MMM d'),
      end: format(e, 'MMM d'),
      days: differenceInCalendarDays(e, s) + 1,
    });

    const svgPoint = (clientX: number, clientY: number): { x: number; y: number } => {
      const rect = svgRef.current!.getBoundingClientRect();
      return { x: clientX - rect.left, y: clientY - rect.top };
    };

    const applyPointer = (clientX: number, clientY: number): void => {
      const L = latest.current!;
      const bodyEl = bodyScrollRef.current;
      const scrollDelta = bodyEl ? bodyEl.scrollLeft - ds.startScrollLeft : 0;
      const dx = clientX - ds.startClientX + scrollDelta;
      if (Math.abs(clientX - ds.startClientX) > 3 || Math.abs(clientY - ds.startClientY) > 3 || scrollDelta !== 0) {
        suppressClickRef.current = true;
      }

      if (ds.mode === 'link') {
        const p = svgPoint(clientX, clientY);
        const line = linkLineRef.current;
        if (line) {
          line.setAttribute('x2', String(p.x));
          line.setAttribute('y2', String(p.y));
        }
        const r = L.rows[Math.floor(p.y / L.rowH)];
        const targetId = r && r.type === 'task' && r.task && r.task.id !== ds.taskId ? r.task.id : null;
        setLinkTargetId(prev => (prev === targetId ? prev : targetId));
        return;
      }

      if (ds.mode === 'create') {
        const p = svgPoint(clientX, clientY);
        const cur = addDays(L.rangeStart, Math.floor(p.x / L.dayWidth));
        const s = cur < ds.origStartDate ? cur : ds.origStartDate;
        const e = cur < ds.origStartDate ? ds.origStartDate : cur;
        const key = `${s.getTime()}|${e.getTime()}`;
        if (key !== lastKeyRef.current) {
          lastKeyRef.current = key;
          setGhost({ row: ds.row!, start: s, end: e });
        }
        setLabel(rangeLabel(s, e), clientX, clientY);
        return;
      }

      const task = L.taskById.get(ds.taskId);
      const deltaDays = Math.round(dx / L.dayWidth);
      const dir: 1 | -1 = deltaDays > 0 ? 1 : -1;
      let result: { start: Date; end: Date } | null = null;

      if (deltaDays !== 0) {
        if (ds.mode === 'move') {
          const s = snapToWorkingDay(addDays(ds.origStartDate, deltaDays), dir, L.cal);
          const span = differenceInCalendarDays(ds.origEndDate, ds.origStartDate);
          let e = addDays(s, span);
          if (span > 0 && !isWorkingDay(dateToDateOnlyString(e), L.cal)) {
            e = snapToWorkingDay(e, -1, L.cal);
            if (e < s) e = s;
          }
          result = { start: s, end: e };
        } else if (ds.mode === 'resize-end') {
          let e = snapToWorkingDay(addDays(ds.origEndDate, deltaDays), dir, L.cal);
          if (e < ds.origStartDate) e = ds.origStartDate;
          if (e.getTime() !== ds.origEndDate.getTime()) result = { start: ds.origStartDate, end: e };
        } else {
          let s = snapToWorkingDay(addDays(ds.origStartDate, deltaDays), dir, L.cal);
          if (s > ds.origEndDate) s = ds.origEndDate;
          if (s.getTime() !== ds.origStartDate.getTime()) result = { start: s, end: ds.origEndDate };
        }
      }

      if (result && result.start.getTime() === ds.origStartDate.getTime() && result.end.getTime() === ds.origEndDate.getTime()) {
        result = null;
      }
      const key = result ? `${result.start.getTime()}|${result.end.getTime()}` : 'none';
      if (key !== lastKeyRef.current) {
        lastKeyRef.current = key;
        const next = new Map<number, { start: Date; end: Date }>();
        if (result) next.set(ds.taskId, result);
        dragOffsetsRef.current = next;
        setDragOffsets(next);
      }
      const shown = result || { start: ds.origStartDate, end: ds.origEndDate };
      setLabel(task && task.isMilestone ? format(shown.start, 'MMM d, yyyy') : rangeLabel(shown.start, shown.end), clientX, clientY);
    };

    const finish = (commit: boolean): void => {
      const L = latest.current!;
      if (commit) {
        if (ds.mode === 'link') {
          const p = svgRef.current ? svgPoint(lastX, lastY) : null;
          if (p && L.onAddDependency) {
            const r = L.rows[Math.floor(p.y / L.rowH)];
            if (r && r.type === 'task' && r.task && r.task.id !== ds.taskId) {
              // The dragged-from bar is the predecessor of the bar it lands on.
              L.onAddDependency(ds.taskId, r.task.id);
            }
          }
          suppressClickRef.current = true;
        } else if (ds.mode === 'create') {
          const moved = Math.abs(lastX - ds.startClientX) > 3 || Math.abs(lastY - ds.startClientY) > 3;
          const g = lastGhostRef.current;
          if (moved && g && L.onCreateTaskAt) {
            L.onCreateTaskAt(dateToDateOnlyString(g.start), dateToDateOnlyString(g.end));
          }
        } else {
          const offset = dragOffsetsRef.current.get(ds.taskId);
          if (offset) {
            suppressClickRef.current = true;
            L.onTaskUpdate(ds.taskId, {
              startDate: dateToDateOnlyString(offset.start),
              dueDate: dateToDateOnlyString(offset.end),
            }, { cascade: true });
          }
        }
      }
      setDragState(null);
      dragOffsetsRef.current = new Map();
      setDragOffsets(new Map());
      setGhost(null);
      setLinkTargetId(null);
    };

    const onMove = (e: PointerEvent): void => {
      if (e.pointerId !== ds.pointerId) return;
      lastX = e.clientX;
      lastY = e.clientY;
      applyPointer(lastX, lastY);
    };
    const onUp = (e: PointerEvent): void => {
      if (e.pointerId !== ds.pointerId) return;
      finish(true);
    };
    // pointercancel = the browser took the gesture (e.g. touch scroll): abandon it.
    const onCancel = (e: PointerEvent): void => {
      if (e.pointerId !== ds.pointerId) return;
      finish(false);
    };
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') {
        e.preventDefault();
        finish(false);
      }
    };
    const onBlur = (): void => finish(false);

    // Edge auto-scroll while the pointer is held near the timeline's edges.
    const tick = (): void => {
      const el = bodyScrollRef.current;
      if (el) {
        const r = el.getBoundingClientRect();
        const edge = 48;
        let vx = 0;
        let vy = 0;
        if (lastX > r.right - edge) vx = Math.min(28, (lastX - (r.right - edge)) / 2 + 2);
        else if (lastX < r.left + edge) vx = -Math.min(28, ((r.left + edge) - lastX) / 2 + 2);
        if (ds.mode === 'link' || ds.mode === 'create') {
          if (lastY > r.bottom - edge) vy = Math.min(28, (lastY - (r.bottom - edge)) / 2 + 2);
          else if (lastY < r.top + edge) vy = -Math.min(28, ((r.top + edge) - lastY) / 2 + 2);
        }
        if (vx || vy) {
          const bx = el.scrollLeft;
          const by = el.scrollTop;
          el.scrollLeft += vx;
          el.scrollTop += vy;
          if (el.scrollLeft !== bx || el.scrollTop !== by) applyPointer(lastX, lastY);
        }
      }
      raf = requestAnimationFrame(tick);
    };
    raf = requestAnimationFrame(tick);
    // Initial label so the readout appears as soon as the drag starts.
    if (ds.mode !== 'link') {
      const startShown = ds.mode === 'create' ? ds.origStartDate : ds.origStartDate;
      setLabel(rangeLabel(startShown, ds.origEndDate), ds.startClientX, ds.startClientY);
    }

    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
    window.addEventListener('pointercancel', onCancel);
    window.addEventListener('keydown', onKey);
    window.addEventListener('blur', onBlur);
    return () => {
      cancelAnimationFrame(raf);
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      window.removeEventListener('pointercancel', onCancel);
      window.removeEventListener('keydown', onKey);
      window.removeEventListener('blur', onBlur);
    };
  }, [dragState]); // eslint-disable-line react-hooks/exhaustive-deps

  lastGhostRef.current = ghost;

  // ─── Header generation ─────────────────────────────────────────────────

  const monthBands = React.useMemo(() => {
    const bands: { label: string; x: number; width: number }[] = [];
    let cur = startOfMonth(rangeStart);
    while (cur <= rangeEnd) {
      const mStart = max([cur, rangeStart]);
      const mEnd = min([endOfMonth(cur), rangeEnd]);
      bands.push({
        label: format(cur, effZoom === 'quarter' ? 'MMM yy' : 'MMMM yyyy'),
        x: differenceInCalendarDays(mStart, rangeStart) * dayWidth,
        width: (differenceInCalendarDays(mEnd, mStart) + 1) * dayWidth,
      });
      cur = addMonths(cur, 1);
    }
    return bands;
  }, [rangeStart, rangeEnd, dayWidth, effZoom]);

  const subBands = React.useMemo(() => {
    if (effZoom === 'day' || effZoom === 'week') {
      // Week bands
      const bands: { label: string; x: number; width: number; isCurrentWeek: boolean }[] = [];
      let cur = startOfWeek(rangeStart, { weekStartsOn: 1 });
      const thisWeekStart = startOfWeek(new Date(todayMs), { weekStartsOn: 1 });
      while (cur <= rangeEnd) {
        const wStart = max([cur, rangeStart]);
        const wEnd = min([addDays(cur, 6), rangeEnd]);
        const w = (differenceInCalendarDays(wEnd, wStart) + 1) * dayWidth;
        bands.push({
          label: getWeekLabel(cur),
          x: differenceInCalendarDays(wStart, rangeStart) * dayWidth,
          width: w,
          isCurrentWeek: cur.getTime() === thisWeekStart.getTime(),
        });
        cur = addWeeks(cur, 1);
      }
      return bands;
    } else {
      // Month bands (for month / quarter zoom)
      return monthBands.map(b => ({ ...b, isCurrentWeek: false }));
    }
  }, [rangeStart, rangeEnd, dayWidth, effZoom, monthBands, settings.weekLabel, projectWeekStart, todayMs]); // eslint-disable-line react-hooks/exhaustive-deps

  // Day tick marks (day zoom only) — only the days near the viewport get
  // elements, so a multi-year range doesn't emit thousands of SVG nodes.
  const dayTicks = React.useMemo(() => {
    if (effZoom !== 'day') return [];
    const ticks: { d: Date; x: number; isToday: boolean; nonWorking: boolean }[] = [];
    const from = Math.max(0, Math.floor(visX0 / dayWidth));
    const to = Math.min(totalDays - 1, Math.ceil(visX1 / dayWidth));
    for (let i = from; i <= to; i++) {
      const d = addDays(rangeStart, i);
      ticks.push({
        d,
        x: i * dayWidth,
        isToday: d.getTime() === todayMs,
        nonWorking: !isWorkingDay(dateToDateOnlyString(d), cal),
      });
    }
    return ticks;
  }, [rangeStart, totalDays, dayWidth, effZoom, visX0, visX1, cal, todayMs]);

  // Non-working shading: one repeating weekly <pattern> for the weekdays plus a
  // rect per holiday, instead of a rect per weekend day across the whole range.
  const showShading = settings.showWeekends && effZoom !== 'quarter';
  const nonWorkingWeekCols = React.useMemo(() => {
    const cols: number[] = [];
    if (cal.workingDays.length === 0) return cols;
    for (let i = 0; i < 7; i++) {
      if (cal.workingDays.indexOf((rangeStart.getDay() + i) % 7) === -1) cols.push(i);
    }
    return cols;
  }, [cal, rangeStart]);
  const holidayCols = React.useMemo(() => {
    const out: number[] = [];
    if (!showShading) return out;
    cal.holidays.forEach(h => {
      const d = parseDateOnly(h);
      if (!d) return;
      const idx = differenceInCalendarDays(d, rangeStart);
      if (idx >= 0 && idx < totalDays) out.push(idx);
    });
    return out;
  }, [cal, rangeStart, totalDays, showShading]);

  // Today x
  const todayX = dateToX(today) + dayWidth / 2;

  // ─── Render helpers ────────────────────────────────────────────────────

  const dragging = !!dragState;

  const renderTaskBars = (): React.ReactNode[] => {
    const out: React.ReactNode[] = [];
    for (let i = rowStart; i < rowEnd; i++) {
      const row = visibleTasks[i];
      if (row.type !== 'task' || !row.task) continue;
      const task = row.task;
      const offset = dragOffsets.get(task.id);
      // A milestone with only a due date set should render there rather than
      // snapping to today — matches the dependency-arrow fallback, which
      // anchors on the same date so arrows don't detach from the diamond.
      const sDate = offset ? offset.start
        : (parseDateOnly(task.startDate) || (task.isMilestone ? parseDateOnly(task.dueDate) : null) || today);
      const eDate = offset ? offset.end : (parseDateOnly(task.dueDate) || today);
      const x = dateToX(sDate);
      const width = Math.max(MIN_BAR_WIDTH, (differenceInCalendarDays(eDate, sDate) + 1) * dayWidth);

      const bs = parseDateOnly(task.baselineStart);
      const be = parseDateOnly(task.baselineDue);

      out.push(
        <GanttBar
          key={`bar-${task.id}`}
          task={task}
          rowIndex={i}
          rowH={ROW_H}
          x={task.isMilestone ? x + dayWidth / 2 : x}
          width={width}
          color={getTaskColor(task, settings)}
          isMilestone={task.isMilestone}
          highlightCritical={settings.showCriticalPath && criticalIds.has(task.id)}
          flat={settings.barStyle === 'flat'}
          showProgressText={settings.showProgressText}
          showAssignee={settings.showAssignee}
          isLinkTarget={linkTargetId === task.id}
          uid={uid}
          hintId={hintId}
          baselineX={bs && be ? dateToX(bs) : undefined}
          baselineWidth={bs && be ? (differenceInCalendarDays(be, bs) + 1) * dayWidth : undefined}
          baselineSlipped={!!(bs && be && eDate > be)}
          handlers={handlers}
        />
      );
    }
    return out;
  };

  const renderDependencyArrows = (): React.ReactNode => {
    const arrows: React.ReactNode[] = [];
    const pad = Math.min(14, Math.max(8, dayWidth * 0.5));

    // Start/end dates of a task as currently drawn (mid-drag offsets included).
    const datesOf = (t: ITask): { s: Date; e: Date } => {
      const off = dragOffsets.get(t.id);
      if (off) return { s: off.start, e: off.end };
      const s = parseDateOnly(t.startDate) || (t.isMilestone ? parseDateOnly(t.dueDate) : null) || today;
      return { s, e: parseDateOnly(t.dueDate) || today };
    };
    // Horizontal anchor of a bar edge; a milestone's edges are its diamond tips.
    const edgeX = (t: ITask, d: { s: Date; e: Date }, side: 'start' | 'end'): number => {
      if (t.isMilestone) return dateToX(d.s) + dayWidth / 2 + (side === 'start' ? -MILESTONE_SIZE : MILESTONE_SIZE);
      if (side === 'start') return dateToX(d.s);
      return Math.max(dateToX(d.s) + MIN_BAR_WIDTH, dateToX(d.e) + dayWidth);
    };

    tasks.forEach(task => {
      if (!task.dependencies.length) return;
      const rowIndex = rowIndexById.get(task.id);
      if (rowIndex === undefined) return;
      task.dependencies.forEach(depId => {
        const dep = taskById.get(depId);
        if (!dep) return;
        const depIndex = rowIndexById.get(depId);
        if (depIndex === undefined) return; // hidden (collapsed phase)
        // Skip arrows entirely outside the rendered row window.
        if ((rowIndex < rowStart && depIndex < rowStart) || (rowIndex >= rowEnd && depIndex >= rowEnd)) return;

        const link = linkOf(task, depId);
        const depDates = datesOf(dep);
        const taskDates = datesOf(task);
        const fromSide = link.type === 'FS' || link.type === 'FF' ? 'end' : 'start';
        const toSide = link.type === 'FS' || link.type === 'SS' ? 'start' : 'end';
        const fdir = fromSide === 'end' ? 1 : -1;
        const tdir = toSide === 'start' ? -1 : 1; // side of the target bar the arrow approaches from
        const fx = edgeX(dep, depDates, fromSide);
        const fy = depIndex * ROW_H + ROW_H / 2;
        const ax = edgeX(task, taskDates, toSide) + tdir * 2;
        const ty = rowIndex * ROW_H + ROW_H / 2;

        const isCritical = criticalIds.has(task.id) && criticalIds.has(depId);
        const beingDragged = dragOffsets.has(task.id) || dragOffsets.has(depId);
        const violated = isLinkViolated(
          { start: depDates.s, end: depDates.e }, { start: taskDates.s, end: taskDates.e }, link, cal
        );

        const isHovered = hoveredId === task.id || hoveredId === depId;
        // Critical path arrows are always visible when showCriticalPathOnly is on.
        // Hovering reveals all arrows regardless of critical path setting.
        // Non-critical arrows are hidden in critical-path-only mode unless hovering.
        // All arrows are hidden when not hovering in hover-only mode.
        // A link that a drag is currently violating is always shown so the
        // warning is visible while the bar is still in the user's hand.
        if (beingDragged && violated) { /* always show */ }
        else if (settings.showCriticalPathOnly && isCritical) { /* always show */ }
        else if (settings.dependenciesOnHover && isHovered) { /* hover reveals all */ }
        else if (settings.showCriticalPathOnly && !isCritical) return;
        else if (settings.dependenciesOnHover && !isHovered) return;

        let pathD: string;
        if (fdir === 1 && tdir === -1 && ax >= fx + pad) {
          // Forward finish-to-start: elbow right→down→right into the bar's start.
          const midX = fx + (ax - fx) / 2;
          pathD = `M ${fx} ${fy} H ${midX} V ${ty} H ${ax}`;
        } else {
          // Every other case: leave the source edge, run along the gap between
          // rows (half a row above/below the source so it never overlaps a
          // bar), then come in horizontally at the target edge.
          const p1 = fx + fdir * pad;
          const q1 = ax + tdir * pad;
          const midY = ty >= fy ? fy + ROW_H * 0.5 : fy - ROW_H * 0.5;
          pathD = `M ${fx} ${fy} H ${p1} V ${midY} H ${q1} V ${ty} H ${ax}`;
        }

        const warn = violated || isCritical;
        arrows.push(
          <g key={`dep-${dep.id}-${task.id}`} pointerEvents="none">
            <path
              d={pathD}
              className={styles.dependencyArrow}
              style={{
                markerEnd: `url(#${uid}-arrow${warn ? '-crit' : ''})`,
                stroke: warn ? CRITICAL_COLOR : undefined,
                strokeWidth: isCritical ? 2 : undefined,
                strokeDasharray: violated ? '5,3' : undefined,
              }}
            />
            {(link.type !== 'FS' || link.lag !== 0) && (
              <text
                className={styles.depLabel}
                x={ax + tdir * 4}
                y={ty - 4}
                textAnchor={tdir === -1 ? 'end' : 'start'}
                fontSize={9}
                fontFamily="'Segoe UI', sans-serif"
                style={warn ? { fill: CRITICAL_COLOR } : undefined}
              >
                {formatDependencyLink(link)}
              </text>
            )}
          </g>
        );
      });
    });
    return arrows;
  };

  // ─── Left panel rows ───────────────────────────────────────────────────

  const renderLeftRows = (): React.ReactNode[] => {
    const out: React.ReactNode[] = [];
    for (let i = rowStart; i < rowEnd; i++) {
      const row = visibleTasks[i];
      const key = row.type === 'phase' ? `phase-${row.phase}` : `row-${row.task!.id}`;
      out.push(
        <LeftRow
          key={key}
          row={row}
          rowH={ROW_H}
          color={row.task ? getTaskColor(row.task, settings) : ''}
          violated={!!row.task && violationIds.has(row.task.id)}
          collapsed={row.type === 'phase' && collapsedPhases.has(row.phase!)}
          onEditTask={onEditTask}
          onDeleteTask={onDeleteTask}
          onTogglePhase={togglePhase}
        />
      );
    }
    return out;
  };

  // ─── Empty state ───────────────────────────────────────────────────────

  if (tasks.length === 0) {
    return (
      <div className={styles.ganttWrapper}>
        <div className={styles.emptyGantt}>
          <div className={styles.emptyIcon}>📅</div>
          <div className={styles.emptyTitle}>{strings.Gantt_EmptyTitle}</div>
          <div className={styles.emptySubtitle}>
            {strings.Gantt_EmptySubtitle}
          </div>
          <div className={styles.emptyActions}>
            <button className={styles.emptyPrimaryBtn} onClick={onAddTask}>
              {strings.Gantt_AddFirstTask}
            </button>
            {onImport && (
              <button className={styles.emptySecondaryBtn} onClick={onImport}>
                {strings.Gantt_ImportTasksButton}
              </button>
            )}
          </div>
        </div>
      </div>
    );
  }

  const nwPatternId = `${uid}-nw`;

  return (
    <div className={`${styles.ganttWrapper} ${dragging ? styles.isDragging : ''}`}>
      {/* Project title bar */}
      <div className={styles.ganttTitleBar} style={{ background: project.color }}>
        <div className={styles.ganttTitleDot} style={{ background: 'rgba(255,255,255,0.3)' }} />
        <span className={styles.ganttTitleText}>{project.title}</span>
        {project.status && (
          <span className={styles.ganttTitleBadge}>{project.status}</span>
        )}
      </div>

      <div className={styles.ganttInner}>
        {/* Left panel */}
        <div className={styles.leftPanel}>
          <div className={styles.leftHeader} style={{ background: theme.bg }}>
            <div className={`${styles.leftHeaderCell} ${styles.taskNameCol}`} style={{ color: theme.subtext }}>
              {strings.Gantt_TaskColumnHeader}
              {phaseNames.length > 0 && (
                <span className={styles.headerBtns}>
                  <button
                    className={styles.headerBtn}
                    title={strings.Gantt_CollapseAll}
                    aria-label={strings.Gantt_CollapseAll}
                    onClick={() => updateCollapsed(new Set(phaseNames))}
                  >
                    ⊟
                  </button>
                  <button
                    className={styles.headerBtn}
                    title={strings.Gantt_ExpandAll}
                    aria-label={strings.Gantt_ExpandAll}
                    onClick={() => updateCollapsed(new Set<string>())}
                  >
                    ⊞
                  </button>
                </span>
              )}
            </div>
            <div className={`${styles.leftHeaderCell} ${styles.durationCol}`} style={{ color: theme.subtext }}>
              {strings.Gantt_DurationColumnHeader}
            </div>
          </div>
          <div
            className={styles.leftBody}
            ref={leftBodyRef}
            onScroll={handleLeftScroll}
            style={{ overflowY: 'auto' }}
          >
            <div style={{ height: svgBodyHeight, position: 'relative' }}>
              <div style={{ position: 'absolute', top: rowStart * ROW_H, left: 0, right: 0 }}>
                {renderLeftRows()}
              </div>
            </div>
            <button className={styles.addTaskRowBtn} onClick={onAddTask}>
              {strings.Gantt_AddTaskButton}
            </button>
          </div>
        </div>

        {/* Right timeline panel */}
        <div className={styles.rightPanel}>
          {/* Sticky header */}
          <div className={styles.timelineHeaderScroll} ref={headerScrollRef}>
            <svg width={svgWidth} height={HEADER_HEIGHT} style={{ display: 'block', background: theme.bg }} aria-hidden="true">
              {/* Month row */}
              {monthBands.map((band, i) => (
                <g key={`month-${i}`}>
                  <line
                    x1={band.x}
                    y1={0}
                    x2={band.x}
                    y2={28}
                    stroke="rgba(255,255,255,0.15)"
                    strokeWidth={1}
                  />
                  <text
                    x={band.x + 8}
                    y={18}
                    fontSize={12}
                    fontWeight="600"
                    fontFamily="'Segoe UI', sans-serif"
                    fill={theme.text}
                    style={{ userSelect: 'none' }}
                  >
                    {band.label}
                  </text>
                </g>
              ))}

              {/* Sub-header row (weeks or days) */}
              {effZoom === 'day' ? (
                dayTicks.map(tick => (
                  <g key={`day-${tick.x}`}>
                    {tick.nonWorking && (
                      <rect x={tick.x} y={28} width={dayWidth} height={28} fill="rgba(255,255,255,0.04)" />
                    )}
                    <text
                      x={tick.x + dayWidth / 2}
                      y={46}
                      fontSize={10}
                      textAnchor="middle"
                      fontFamily="'Segoe UI', sans-serif"
                      fill={tick.isToday ? '#FFD700' : theme.subtext}
                      fontWeight={tick.isToday ? '700' : '400'}
                      style={{ userSelect: 'none' }}
                    >
                      {format(tick.d, 'd')}
                    </text>
                    {tick.isToday && (
                      <rect x={tick.x + 2} y={28} width={dayWidth - 4} height={26} rx={3} fill="rgba(255,215,0,0.15)" />
                    )}
                  </g>
                ))
              ) : (
                subBands.map((band, i) => (
                  <g key={`sub-${i}`}>
                    <line
                      x1={band.x}
                      y1={28}
                      x2={band.x}
                      y2={56}
                      stroke="rgba(255,255,255,0.1)"
                      strokeWidth={1}
                    />
                    {band.isCurrentWeek && (
                      <rect x={band.x} y={28} width={band.width} height={28} fill="rgba(255,215,0,0.1)" />
                    )}
                    <text
                      x={band.x + 6}
                      y={46}
                      fontSize={10}
                      fontFamily="'Segoe UI', sans-serif"
                      fill={band.isCurrentWeek ? '#FFD700' : theme.subtext}
                      fontWeight={band.isCurrentWeek ? '700' : '400'}
                      style={{ userSelect: 'none' }}
                    >
                      {band.label}
                    </text>
                  </g>
                ))
              )}
            </svg>
          </div>

          {/* Scrollable body */}
          <div
            className={styles.timelineBodyScroll}
            ref={bodyScrollRef}
            onScroll={handleBodyScroll}
          >
            <svg
              ref={svgRef}
              width={svgWidth}
              height={bodyHeight}
              style={{ display: 'block' }}
              role="group"
              aria-label={strings.Gantt_TimelineAriaLabel}
            >
              <defs>
                <marker id={`${uid}-arrow`} markerWidth="4" markerHeight="3" refX="3.5" refY="1.5" orient="auto">
                  <polygon className={styles.arrowHead} points="0 0, 4 1.5, 0 3" />
                </marker>
                <marker id={`${uid}-arrow-crit`} markerWidth="4" markerHeight="3" refX="3.5" refY="1.5" orient="auto">
                  <polygon className={styles.arrowHeadCrit} points="0 0, 4 1.5, 0 3" />
                </marker>
                {settings.barStyle !== 'flat' && distinctBarColors.map(c => (
                  <linearGradient key={c} id={`${uid}-grad-${colorId(c)}`} x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor={c} stopOpacity="1" />
                    <stop offset="100%" stopColor={c} stopOpacity="0.75" />
                  </linearGradient>
                ))}
                {showShading && nonWorkingWeekCols.length > 0 && (
                  <pattern id={nwPatternId} patternUnits="userSpaceOnUse" x={0} y={0} width={7 * dayWidth} height={64}>
                    {nonWorkingWeekCols.map(i => (
                      <rect key={i} className={styles.nonWorking} x={i * dayWidth} y={0} width={dayWidth} height={64} />
                    ))}
                  </pattern>
                )}
              </defs>

              {/* Non-working day shading (weekends per the working calendar + holidays) */}
              {showShading && nonWorkingWeekCols.length > 0 && (
                <rect x={0} y={0} width={svgWidth} height={bodyHeight} fill={`url(#${nwPatternId})`} pointerEvents="none" />
              )}
              {holidayCols.map(idx => (
                <rect
                  key={`hol-${idx}`}
                  className={styles.nonWorking}
                  x={idx * dayWidth}
                  y={0}
                  width={dayWidth}
                  height={bodyHeight}
                  pointerEvents="none"
                />
              ))}

              {/* Row backgrounds + horizontal lines (visible window only) */}
              {visibleTasks.slice(rowStart, rowEnd).map((row, k) => {
                const i = rowStart + k;
                return (
                  <g key={`rowbg-${i}`} pointerEvents="none">
                    {row.type === 'phase' && (
                      <rect className={styles.phaseBand} x={0} y={i * ROW_H} width={svgWidth} height={ROW_H} />
                    )}
                    <line
                      className={styles.rowLine}
                      x1={0}
                      y1={(i + 1) * ROW_H}
                      x2={svgWidth}
                      y2={(i + 1) * ROW_H}
                    />
                  </g>
                );
              })}

              {/* Today highlight column */}
              <rect
                className={styles.todayColumn}
                x={dateToX(today)}
                y={0}
                width={dayWidth}
                height={bodyHeight}
                pointerEvents="none"
              />

              {/* Hit area for drag-to-create on empty rows (below arrows + bars) */}
              <rect
                x={0}
                y={0}
                width={svgWidth}
                height={svgBodyHeight}
                fill="transparent"
                onPointerDown={handleBackgroundPointerDown}
              />

              {/* Ghost of the task being drawn */}
              {ghost && (
                <rect
                  className={styles.ghostBar}
                  x={dateToX(ghost.start)}
                  y={ghost.row * ROW_H + (ROW_H - BAR_HEIGHT) / 2}
                  width={(differenceInCalendarDays(ghost.end, ghost.start) + 1) * dayWidth}
                  height={BAR_HEIGHT}
                  rx={4}
                  pointerEvents="none"
                />
              )}

              {/* Dependency arrows behind the bars */}
              {settings.showDependencies && renderDependencyArrows()}

              {/* Task bars */}
              {renderTaskBars()}

              {/* Rubber-band line while drawing a dependency */}
              {dragState && dragState.mode === 'link' && (
                <line
                  ref={linkLineRef}
                  className={styles.linkLine}
                  x1={dragState.fromX}
                  y1={dragState.fromY}
                  x2={dragState.fromX}
                  y2={dragState.fromY}
                  pointerEvents="none"
                />
              )}

              {/* Today line */}
              <line
                className={styles.todayLineSvg}
                x1={todayX}
                y1={0}
                x2={todayX}
                y2={bodyHeight}
                strokeWidth={2}
                strokeDasharray="4,3"
              />
              <circle className={styles.todayDot} cx={todayX} cy={4} r={4} />
            </svg>
          </div>
        </div>
      </div>

      {/* Hover/focus tooltip — owns its own state so bars don't re-render on hover */}
      <GanttTooltip apiRef={tooltipApiRef} showHealthBadges={settings.showHealthBadges} />

      {/* Live date readout while dragging */}
      {dragging && dragState!.mode !== 'link' && <div className={styles.dragLabel} ref={dragLabelRef} aria-hidden="true" />}

      {/* Screen-reader support: keyboard instructions + announcements */}
      <div id={hintId} className={styles.srOnly}>{strings.Gantt_KeyboardHint}</div>
      <div className={styles.srOnly} role="status" aria-live="polite">{announce}</div>
    </div>
  );
};

export default GanttChart;
