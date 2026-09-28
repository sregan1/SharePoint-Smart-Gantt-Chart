import * as React from 'react';
import {
  addDays, addMonths, addWeeks, differenceInCalendarDays,
  endOfMonth, format, isWeekend, max, min,
  startOfMonth, startOfWeek, getISOWeek,
} from 'date-fns';
import {
  IProject, ITask, ZoomLevel, STATUS_COLORS, PRIORITY_COLORS,
  IGanttDisplaySettings, DEFAULT_GANTT_SETTINGS, HEADER_THEME_COLORS, phaseColor,
} from '../../models';
import {
  computeTaskHealth, healthColor, computeCriticalPath, hasDependencyViolation,
} from '../../utils/healthUtils';
import { parseDateOnly, formatDateOnly, dateToDateOnlyString, todayLocalMidnight } from '../../utils/dateUtils';
import { HealthBadge } from '../common/HealthBadge';
import * as strings from 'SmartGanttWebPartStrings';
import { formatString } from '../localeUtils';
import styles from './GanttChart.module.scss';

interface IGanttChartProps {
  tasks: ITask[];
  project: IProject;
  zoomLevel: ZoomLevel;
  settings: IGanttDisplaySettings;
  scrollToToday: boolean;
  onEditTask: (task: ITask) => void;
  onDeleteTask: (id: number) => void;
  onTaskUpdate: (id: number, updates: Partial<ITask>) => void;
  onAddTask: () => void;
  onImport?: () => void;
}

type DragMode = 'move' | 'resize-start' | 'resize-end';

interface IDragState {
  taskId: number;
  mode: DragMode;
  startClientX: number;
  origStartDate: Date;
  origEndDate: Date;
}

interface ITooltip {
  x: number;
  y: number;
  task: ITask;
}

const HEADER_HEIGHT = 56;
const BAR_HEIGHT = 26;
const MILESTONE_SIZE = 12;
const MIN_BAR_WIDTH = 6;

const DAY_WIDTH: Record<ZoomLevel, number> = {
  day: 42,
  week: 18,
  month: 7,
  quarter: 4,
};

const BUFFER_DAYS = 30;

function taskDuration(task: ITask): number {
  const s = parseDateOnly(task.startDate);
  const e = parseDateOnly(task.dueDate);
  if (!s || !e) return 0;
  return Math.max(1, differenceInCalendarDays(e, s) + 1);
}

const HEX_COLOR_RE = /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/;

function getTaskColor(task: ITask, settings: IGanttDisplaySettings): string {
  // task.color is free-form user input (models/index.ts) — a stray value
  // (short hex, a CSS name, import garbage) must not flow into SVG color
  // attributes unvalidated, or the bar renders with NaN-derived colors.
  if (task.color && HEX_COLOR_RE.test(task.color)) return task.color;
  if (settings.colorBy === 'priority') return PRIORITY_COLORS[task.priority] || '#0078D4';
  if (settings.colorBy === 'phase' && task.phase) return phaseColor(task.phase);
  if (settings.colorBy === 'health') return healthColor(computeTaskHealth(task));
  return STATUS_COLORS[task.status] || '#0078D4';
}

function hexToRgba(hex: string, alpha: number): string {
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
function colorId(hex: string): string {
  return hex.replace('#', '');
}

let ganttInstanceCounter = 0;

export const GanttChart: React.FC<IGanttChartProps> = ({
  tasks,
  project,
  zoomLevel,
  settings = DEFAULT_GANTT_SETTINGS,
  scrollToToday,
  onEditTask,
  onDeleteTask,
  onTaskUpdate,
  onAddTask,
  onImport,
}) => {
  const bodyScrollRef = React.useRef<HTMLDivElement>(null);
  const headerScrollRef = React.useRef<HTMLDivElement>(null);
  const leftBodyRef = React.useRef<HTMLDivElement>(null);

  // Unique per-instance prefix so SVG defs (gradients, markers) don't collide
  // when two Smart Gantt web parts render on the same page. Lazily
  // initialized so the module counter increments once per mount, not once
  // per render (React.useRef(expr) evaluates expr on every render even
  // though only the first value is kept).
  const uidRef = React.useRef<string>();
  if (!uidRef.current) uidRef.current = `sg${++ganttInstanceCounter}`;
  const uid = uidRef.current;

  const [dragState, setDragState] = React.useState<IDragState | null>(null);
  const [dragOffsets, setDragOffsets] = React.useState<Map<number, { start: Date; end: Date }>>(new Map());
  // Live mirror of dragOffsets so the mousemove/mouseup listeners don't need
  // re-registering on every state change.
  const dragOffsetsRef = React.useRef(dragOffsets);
  // Set after a real drag so the click that the browser fires on mouseup
  // doesn't also open the edit panel.
  const suppressClickRef = React.useRef(false);
  const [tooltip, setTooltip] = React.useState<ITooltip | null>(null);
  const [collapsedPhases, setCollapsedPhases] = React.useState<Set<string>>(new Set());
  // Tracks whether the previous render had scrollToToday=true, so the
  // false-reset 100ms later doesn't override the Today scroll with the
  // smart-initial-position logic.
  const prevScrollToToday = React.useRef(false);

  const dayWidth = DAY_WIDTH[zoomLevel];
  const today = todayLocalMidnight();
  const ROW_H = settings.rowHeight;
  const theme = HEADER_THEME_COLORS[settings.headerTheme];

  // For project-relative week numbers: find earliest task start
  const projectWeekStart = React.useMemo(() => {
    const earliest = tasks.reduce<Date | null>((acc, t) => {
      const s = parseDateOnly(t.startDate);
      return s && (!acc || s < acc) ? s : acc;
    }, null);
    return startOfWeek(earliest || today, { weekStartsOn: 1 });
  }, [tasks]);

  const criticalIds = React.useMemo(
    () => (settings.showCriticalPath || settings.showCriticalPathOnly ? computeCriticalPath(tasks) : new Set<number>()),
    [tasks, settings.showCriticalPath, settings.showCriticalPathOnly]
  );

  const violationIds = React.useMemo(() => {
    const byId = new Map(tasks.map(t => [t.id, t]));
    const set = new Set<number>();
    tasks.forEach(t => { if (hasDependencyViolation(t, byId)) set.add(t.id); });
    return set;
  }, [tasks]);

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
    const dates: Date[] = [addDays(today, -BUFFER_DAYS)];
    tasks.forEach(t => {
      const s = parseDateOnly(t.startDate);
      const e = parseDateOnly(t.dueDate);
      if (s) dates.push(addDays(s, -BUFFER_DAYS));
      if (e) dates.push(addDays(e, BUFFER_DAYS));
    });
    const earliest = startOfMonth(dates.reduce((a, b) => (a < b ? a : b), today));
    const latest = endOfMonth(dates.reduce((a, b) => (a > b ? a : b), today));
    return { rangeStart: earliest, rangeEnd: latest };
  }, [tasks, today.getTime()]);

  const totalDays = differenceInCalendarDays(rangeEnd, rangeStart) + 1;
  const svgWidth = totalDays * dayWidth;
  // Rebuilding this per render is otherwise the most expensive step in the
  // component — it runs on every mousemove during a drag and every tooltip
  // hover, neither of which change tasks/collapsedPhases.
  const visibleTasks = React.useMemo(
    () => buildVisibleRows(tasks, collapsedPhases),
    [tasks, collapsedPhases]
  );
  const svgBodyHeight = visibleTasks.length * ROW_H;

  // Shared by renderDependencyArrows on every render (including drag/hover
  // renders that don't change tasks/visibleTasks) — hoisted so those renders
  // don't rebuild both maps from scratch each time.
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

  // Convert date ↔ x
  const dateToX = (d: Date): number =>
    differenceInCalendarDays(d, rangeStart) * dayWidth;

  // Scroll on load, zoom change, or ◉ Today button.
  // On load/range change: scroll to the earliest task so past tasks are visible.
  // On Today button: always snap to today regardless of task dates.
  // prevScrollToToday guards against the boolean false-reset 100ms after Today
  // fires overriding the Today scroll with the smart-position logic.
  React.useEffect(() => {
    if (!bodyScrollRef.current) return;

    const isToday = scrollToToday || prevScrollToToday.current;
    prevScrollToToday.current = scrollToToday;

    let scrollX: number;
    if (isToday) {
      scrollX = Math.max(0, dateToX(today) - 200);
    } else {
      // Find the earliest task start date
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

    bodyScrollRef.current.scrollLeft = scrollX;
    if (headerScrollRef.current) headerScrollRef.current.scrollLeft = scrollX;
    // rangeStart is a new Date object every render the task list changes
    // (even when the *value* is unchanged — e.g. editing an unrelated field
    // triggers a task refetch), so depending on it directly would re-snap
    // the viewport away from wherever the user scrolled after every edit.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [scrollToToday, rangeStart.getTime(), dayWidth]);

  // Sync scroll between panels
  const handleBodyScroll = (): void => {
    if (!bodyScrollRef.current) return;
    if (headerScrollRef.current)
      headerScrollRef.current.scrollLeft = bodyScrollRef.current.scrollLeft;
    if (leftBodyRef.current)
      leftBodyRef.current.scrollTop = bodyScrollRef.current.scrollTop;
  };

  const handleLeftScroll = (): void => {
    if (!leftBodyRef.current || !bodyScrollRef.current) return;
    bodyScrollRef.current.scrollTop = leftBodyRef.current.scrollTop;
  };

  // ─── Drag handlers ─────────────────────────────────────────────────────
  // Pointer Events (rather than mouse-only) so dragging/resizing bars also
  // works with touch and pen input, which SharePoint pages are commonly
  // viewed on (tablets). The window-level pointermove/pointerup listeners
  // below still receive events normally — capture isn't needed since the
  // events keep bubbling to window regardless of which element they target.

  const handleBarPointerDown = (
    e: React.PointerEvent,
    task: ITask,
    mode: DragMode
  ): void => {
    e.preventDefault();
    e.stopPropagation();
    suppressClickRef.current = false;
    setTooltip(null);
    const s = parseDateOnly(task.startDate) || today;
    const end = parseDateOnly(task.dueDate) || today;
    setDragState({
      taskId: task.id,
      mode,
      startClientX: e.clientX,
      origStartDate: s,
      origEndDate: end,
    });
  };

  // Last deltaDays actually applied during the current drag — a day is
  // several pixels wide, so most raw pointermove events round to the same
  // delta and would otherwise re-render the whole chart (bars, arrows, left
  // panel) for no visible change.
  const lastDeltaRef = React.useRef<number | null>(null);

  React.useEffect(() => {
    if (!dragState) return;
    lastDeltaRef.current = null;

    const handlePointerMove = (e: PointerEvent): void => {
      const dx = e.clientX - dragState.startClientX;
      if (Math.abs(dx) > 3) suppressClickRef.current = true;
      const deltaDays = Math.round(dx / dayWidth);
      if (deltaDays === lastDeltaRef.current) return;
      lastDeltaRef.current = deltaDays;
      const newMap = new Map(dragOffsetsRef.current);

      if (deltaDays === 0) {
        newMap.delete(dragState.taskId);
      } else if (dragState.mode === 'move') {
        newMap.set(dragState.taskId, {
          start: addDays(dragState.origStartDate, deltaDays),
          end: addDays(dragState.origEndDate, deltaDays),
        });
      } else if (dragState.mode === 'resize-end') {
        const newEnd = addDays(dragState.origEndDate, deltaDays);
        if (differenceInCalendarDays(newEnd, dragState.origStartDate) >= 0) {
          newMap.set(dragState.taskId, {
            start: dragState.origStartDate,
            end: newEnd,
          });
        }
      } else {
        // resize-start
        const newStart = addDays(dragState.origStartDate, deltaDays);
        if (differenceInCalendarDays(dragState.origEndDate, newStart) >= 0) {
          newMap.set(dragState.taskId, {
            start: newStart,
            end: dragState.origEndDate,
          });
        }
      }
      dragOffsetsRef.current = newMap;
      setDragOffsets(newMap);
    };

    const handlePointerUp = (): void => {
      const offset = dragOffsetsRef.current.get(dragState.taskId);
      if (offset) {
        suppressClickRef.current = true;
        onTaskUpdate(dragState.taskId, {
          startDate: dateToDateOnlyString(offset.start),
          dueDate: dateToDateOnlyString(offset.end),
        });
      }
      setDragState(null);
      dragOffsetsRef.current = new Map();
      setDragOffsets(new Map());
    };

    window.addEventListener('pointermove', handlePointerMove);
    window.addEventListener('pointerup', handlePointerUp);
    window.addEventListener('pointercancel', handlePointerUp);
    return () => {
      window.removeEventListener('pointermove', handlePointerMove);
      window.removeEventListener('pointerup', handlePointerUp);
      window.removeEventListener('pointercancel', handlePointerUp);
    };
  }, [dragState, dayWidth]); // eslint-disable-line react-hooks/exhaustive-deps

  const handleBarClick = (task: ITask): void => {
    if (suppressClickRef.current) {
      suppressClickRef.current = false;
      return;
    }
    onEditTask(task);
  };

  // ─── Header generation ─────────────────────────────────────────────────

  const monthBands = React.useMemo(() => {
    const bands: { label: string; x: number; width: number }[] = [];
    let cur = startOfMonth(rangeStart);
    while (cur <= rangeEnd) {
      const mStart = max([cur, rangeStart]);
      const mEnd = min([endOfMonth(cur), rangeEnd]);
      bands.push({
        label: format(cur, zoomLevel === 'quarter' ? 'MMM yy' : 'MMMM yyyy'),
        x: dateToX(mStart),
        width: (differenceInCalendarDays(mEnd, mStart) + 1) * dayWidth,
      });
      cur = addMonths(cur, 1);
    }
    return bands;
  }, [rangeStart, rangeEnd, dayWidth, zoomLevel]);

  const subBands = React.useMemo(() => {
    if (zoomLevel === 'day' || zoomLevel === 'week') {
      // Week bands
      const bands: { label: string; x: number; width: number; isCurrentWeek: boolean }[] = [];
      let cur = startOfWeek(rangeStart, { weekStartsOn: 1 });
      const thisWeekStart = startOfWeek(today, { weekStartsOn: 1 });
      while (cur <= rangeEnd) {
        const wStart = max([cur, rangeStart]);
        const wEnd = min([addDays(cur, 6), rangeEnd]);
        const w = (differenceInCalendarDays(wEnd, wStart) + 1) * dayWidth;
        bands.push({
          label: getWeekLabel(cur),
          x: dateToX(wStart),
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
  }, [rangeStart, rangeEnd, dayWidth, zoomLevel, monthBands, settings.weekLabel, projectWeekStart]); // eslint-disable-line react-hooks/exhaustive-deps

  // Day tick marks (only in day zoom)
  const dayTicks = React.useMemo(() => {
    if (zoomLevel !== 'day') return [];
    const ticks: { d: Date; x: number; isToday: boolean; isWeekend: boolean }[] = [];
    let cur = rangeStart;
    while (cur <= rangeEnd) {
      ticks.push({
        d: cur,
        x: dateToX(cur),
        isToday: cur.getTime() === today.getTime(),
        isWeekend: isWeekend(cur),
      });
      cur = addDays(cur, 1);
    }
    return ticks;
  }, [rangeStart, rangeEnd, dayWidth, zoomLevel]); // eslint-disable-line react-hooks/exhaustive-deps

  // Weekend columns (for body)
  const weekendCols = React.useMemo(() => {
    if (zoomLevel === 'quarter' || !settings.showWeekends) return [];
    const cols: { x: number; width: number }[] = [];
    let cur = rangeStart;
    while (cur <= rangeEnd) {
      if (isWeekend(cur)) {
        cols.push({ x: dateToX(cur), width: dayWidth });
      }
      cur = addDays(cur, 1);
    }
    return cols;
  }, [rangeStart, rangeEnd, dayWidth, zoomLevel, settings.showWeekends]); // eslint-disable-line react-hooks/exhaustive-deps

  // Today x
  const todayX = dateToX(today) + dayWidth / 2;

  // ─── Render helpers ────────────────────────────────────────────────────

  const BAR_OFFSET = (ROW_H - BAR_HEIGHT) / 2;

  const renderTaskBar = (task: ITask, rowIndex: number): React.ReactNode => {
    const offset = dragOffsets.get(task.id);
    // A milestone with only a due date set should render there rather than
    // snapping to today — matches the dependency-arrow fallback below, which
    // anchors on the same date so arrows don't detach from the diamond.
    const sDate = offset ? offset.start
      : (parseDateOnly(task.startDate) || (task.isMilestone ? parseDateOnly(task.dueDate) : null) || today);
    const eDate = offset ? offset.end : (parseDateOnly(task.dueDate) || today);

    const x = dateToX(sDate);
    const barWidth = Math.max(
      MIN_BAR_WIDTH,
      (differenceInCalendarDays(eDate, sDate) + 1) * dayWidth
    );
    const y = rowIndex * ROW_H + BAR_OFFSET;
    const color = getTaskColor(task, settings);
    const progressWidth = barWidth * (task.percentComplete / 100);
    // Raw critical-path membership — used for the dependency-arrow color/width
    // logic regardless of the display toggle below.
    const isCritical = criticalIds.has(task.id);
    // Bar/milestone decoration only shows when the user has actually turned
    // on "Critical path highlight"; criticalIds itself is also populated for
    // showCriticalPathOnly (arrow visibility), which must not force the
    // outline on when that display toggle is off.
    const highlightCritical = settings.showCriticalPath && isCritical;

    if (task.isMilestone) {
      const mx = dateToX(sDate) + dayWidth / 2;
      const my = rowIndex * ROW_H + ROW_H / 2;
      const milestoneAriaLabel = formatString(strings.Gantt_MilestoneAriaLabel, {
        taskTitle: task.title,
        date: formatDateOnly(dateToDateOnlyString(sDate), 'MMM d, yyyy'),
      });
      return (
        <g
          key={`bar-${task.id}`}
          className={styles.taskBarGroup}
          role="img"
          aria-label={milestoneAriaLabel}
          onClick={() => onEditTask(task)}
          onMouseEnter={e => setTooltip({ x: e.clientX, y: e.clientY, task })}
          onMouseLeave={() => setTooltip(null)}
        >
          <polygon
            points={`${mx},${my - MILESTONE_SIZE} ${mx + MILESTONE_SIZE},${my} ${mx},${my + MILESTONE_SIZE} ${mx - MILESTONE_SIZE},${my}`}
            fill={color}
            stroke={highlightCritical ? '#D13438' : 'white'}
            strokeWidth="1.5"
          />
        </g>
      );
    }

    // "Flat" bar style fills with the plain color; "Gradient" (default)
    // references the shared per-color gradient hoisted into the SVG's <defs>.
    const progressFill = settings.barStyle === 'flat' ? color : `url(#${uid}-grad-${colorId(color)})`;
    const barAriaLabel = formatString(strings.Gantt_BarAriaLabel, {
      taskTitle: task.title,
      startDate: formatDateOnly(task.startDate, 'MMM d, yyyy'),
      dueDate: formatDateOnly(task.dueDate, 'MMM d, yyyy'),
      percent: task.percentComplete,
    });
    return (
      <g key={`bar-${task.id}`} className={styles.taskBarGroup} role="img" aria-label={barAriaLabel}>
        {/* Background bar */}
        <rect
          x={x}
          y={y}
          width={barWidth}
          height={BAR_HEIGHT}
          rx={4}
          fill={hexToRgba(color, 0.18)}
          stroke={highlightCritical ? '#D13438' : undefined}
          strokeWidth={highlightCritical ? 1.5 : undefined}
          strokeDasharray={highlightCritical ? '4,2' : undefined}
          className={styles.taskBar}
          onPointerDown={e => handleBarPointerDown(e, task, 'move')}
          onClick={() => handleBarClick(task)}
          onMouseEnter={e => setTooltip({ x: e.clientX, y: e.clientY, task })}
          onMouseLeave={() => setTooltip(null)}
          style={{ cursor: dragState ? 'grabbing' : 'grab', touchAction: 'none' }}
        />
        {/* Progress fill */}
        {progressWidth > 0 && (
          <rect
            x={x}
            y={y}
            width={Math.min(progressWidth, barWidth)}
            height={BAR_HEIGHT}
            rx={4}
            fill={progressFill}
            pointerEvents="none"
          />
        )}
        {/* Progress label */}
        {settings.showProgressText && barWidth > 50 && task.percentComplete > 0 && (
          <text
            x={x + 8}
            y={y + BAR_HEIGHT / 2 + 4}
            fontSize={11}
            fontFamily="'Segoe UI', sans-serif"
            fontWeight="600"
            fill={progressWidth > barWidth * 0.4 ? '#ffffff' : color}
            pointerEvents="none"
            style={{ userSelect: 'none' }}
          >
            {`${task.percentComplete}%`}
          </text>
        )}
        {/* Assignee label next to the bar */}
        {settings.showAssignee && task.assignedTo && (
          <text
            x={x + barWidth + 8}
            y={y + BAR_HEIGHT / 2 + 4}
            fontSize={11}
            fontFamily="'Segoe UI', sans-serif"
            fill="#605E5C"
            pointerEvents="none"
            style={{ userSelect: 'none' }}
          >
            {task.assignedTo}
          </text>
        )}
        {/* Resize handles */}
        <rect
          x={x}
          y={y}
          width={8}
          height={BAR_HEIGHT}
          rx={4}
          fill={color}
          opacity={0.5}
          className={styles.taskBarResizeHandle}
          onPointerDown={e => handleBarPointerDown(e, task, 'resize-start')}
          style={{ cursor: 'ew-resize', touchAction: 'none' }}
        />
        <rect
          x={x + barWidth - 8}
          y={y}
          width={8}
          height={BAR_HEIGHT}
          rx={4}
          fill={color}
          opacity={0.5}
          className={styles.taskBarResizeHandle}
          onPointerDown={e => handleBarPointerDown(e, task, 'resize-end')}
          style={{ cursor: 'ew-resize', touchAction: 'none' }}
        />
      </g>
    );
  };

  const renderDependencyArrows = (): React.ReactNode => {
    const arrows: React.ReactNode[] = [];
    visibleTasks.forEach((row, rowIndex) => {
      if (row.type !== 'task' || !row.task) return;
      const task = row.task;
      task.dependencies.forEach(depId => {
        const dep = taskById.get(depId);
        if (!dep) return;
        const depIndex = rowIndexById.get(depId);
        if (depIndex === undefined) return; // hidden (collapsed phase)

        const offset = dragOffsets.get(dep.id);
        const depEnd = offset ? offset.end : (parseDateOnly(dep.dueDate) || today);
        const offset2 = dragOffsets.get(task.id);
        const taskStart = offset2 ? offset2.start : (parseDateOnly(task.startDate) || today);

        const toX = dateToX(taskStart) - 2;
        const toY = rowIndex * ROW_H + ROW_H / 2;
        const isCritical = criticalIds.has(task.id) && criticalIds.has(depId);

        const isHovered = tooltip?.task?.id === task.id || tooltip?.task?.id === depId;
        // Critical path arrows are always visible when showCriticalPathOnly is on.
        // Hovering reveals all arrows regardless of critical path setting.
        // Non-critical arrows are hidden in critical-path-only mode unless hovering.
        // All arrows are hidden when not hovering in hover-only mode.
        if (settings.showCriticalPathOnly && isCritical) { /* always show */ }
        else if (settings.dependenciesOnHover && isHovered) { /* hover reveals all */ }
        else if (settings.showCriticalPathOnly && !isCritical) return;
        else if (settings.dependenciesOnHover && !isHovered) return;

        const routePad = Math.max(dayWidth * 0.6, 10);
        let pathD: string;
        if (dep.isMilestone) {
          // Use startDate so mx matches where the diamond is actually rendered (renderBar uses startDate).
          const milestoneDate = offset ? offset.start : (parseDateOnly(dep.startDate) || depEnd);
          const mx = dateToX(milestoneDate) + dayWidth / 2;
          const my = depIndex * ROW_H + ROW_H / 2;
          const myBottom = my + MILESTONE_SIZE;
          if (toX > mx) {
            // Target is to the right: straight down from bottom tip, then right into middle-left of bar.
            pathD = `M ${mx} ${myBottom} V ${toY} H ${toX}`;
          } else {
            // Same-date or backward: drop straight down into the top of the bar (arrowhead points ↓).
            const barTopY = rowIndex * ROW_H + (ROW_H - BAR_HEIGHT) / 2 - 2;
            pathD = `M ${mx} ${myBottom} V ${barTopY}`;
          }
        } else {
          const fromX = dateToX(depEnd) + dayWidth;
          const fromY = depIndex * ROW_H + ROW_H / 2;
          if (toX >= fromX) {
            // Forward: elbow right→down→right into middle-left of bar.
            const midX = fromX + (toX - fromX) / 2;
            pathD = `M ${fromX} ${fromY} H ${midX} V ${toY} H ${toX}`;
          } else {
            // Backward: horizontal routing lane sits in the mid-gap between rows (0.5×ROW_H
            // above/below the target centre) so it never overlaps the target bar.
            const overshoot = toY >= fromY ? toY - ROW_H * 0.5 : toY + ROW_H * 0.5;
            pathD = `M ${fromX} ${fromY} H ${fromX + routePad} V ${overshoot} H ${toX - routePad} V ${toY} H ${toX}`;
          }
        }

        arrows.push(
          <path
            key={`dep-${dep.id}-${task.id}`}
            d={pathD}
            className={styles.dependencyArrow}
            style={{
              markerEnd: `url(#${uid}-arrow${isCritical ? '-crit' : ''})`,
              stroke: isCritical ? '#D13438' : undefined,
              strokeWidth: isCritical ? 2 : undefined,
            }}
          />
        );
      });
    });
    return arrows;
  };

  // ─── Left panel rows ───────────────────────────────────────────────────

  const renderLeftRows = (): React.ReactNode => {
    return visibleTasks.map((row, _i) => {
      if (row.type === 'phase') {
        const isCollapsed = collapsedPhases.has(row.phase!);
        return (
          <div
            key={`phase-${row.phase}`}
            className={`${styles.taskRow} ${styles.phaseRow}`}
            style={{ height: ROW_H }}
          >
            <button
              className={styles.taskExpandBtn}
              aria-label={formatString(strings.Gantt_ExpandPhaseAriaLabel, {
                action: isCollapsed ? strings.Gantt_ExpandAction : strings.Gantt_CollapseAction,
                phase: row.phase!,
              })}
              aria-expanded={!isCollapsed}
              onClick={() => {
                setCollapsedPhases(prev => {
                  const next = new Set(prev);
                  isCollapsed ? next.delete(row.phase!) : next.add(row.phase!);
                  return next;
                });
              }}
            >
              {isCollapsed ? '▶' : '▼'}
            </button>
            <span className={styles.phaseLabel}>&ensp;{row.phase}</span>
          </div>
        );
      }

      const task = row.task!;
      const isChild = row.isChild;
      const dur = taskDuration(task);
      const hasViolation = violationIds.has(task.id);
      return (
        <div
          key={`row-${task.id}`}
          className={styles.taskRow}
          style={{ height: ROW_H }}
        >
          {isChild && <div className={styles.taskIndent} />}
          <div
            className={styles.taskStatusDot}
            style={{ background: getTaskColor(task, settings) }}
          />
          {task.isMilestone && <span className={styles.milestoneIcon}>◆</span>}
          <span className={styles.taskName} title={task.title}>{task.title}</span>
          {hasViolation && (
            <span
              title={strings.Gantt_DependencyViolationTooltip}
              style={{ color: '#CA5010', fontSize: 12, flexShrink: 0 }}
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
    });
  };

  // ─── Empty state ───────────────────────────────────────────────────────

  if (tasks.length === 0) {
    return (
      <div className={styles.ganttWrapper}>
        <div className={styles.emptyGantt}>
          <div style={{ fontSize: 48, opacity: 0.3 }}>📅</div>
          <div style={{ fontSize: 18, fontWeight: 600, color: '#323130' }}>{strings.Gantt_EmptyTitle}</div>
          <div style={{ fontSize: 14, color: '#605E5C' }}>
            {strings.Gantt_EmptySubtitle}
          </div>
          <div style={{ display: 'flex', gap: 10 }}>
            <button
              style={{
                background: '#0078D4', color: '#fff', border: 'none', borderRadius: 4,
                padding: '8px 20px', fontSize: 14, fontWeight: 600, cursor: 'pointer',
              }}
              onClick={onAddTask}
            >
              {strings.Gantt_AddFirstTask}
            </button>
            {onImport && (
              <button
                style={{
                  background: '#fff', color: '#0078D4', border: '1px solid #0078D4', borderRadius: 4,
                  padding: '8px 20px', fontSize: 14, fontWeight: 600, cursor: 'pointer',
                }}
                onClick={onImport}
              >
                {strings.Gantt_ImportTasksButton}
              </button>
            )}
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.ganttWrapper}>
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
            <div className={`${styles.leftHeaderCell} ${styles.taskNameCol}`}>{strings.Gantt_TaskColumnHeader}</div>
            <div className={`${styles.leftHeaderCell} ${styles.durationCol}`}>{strings.Gantt_DurationColumnHeader}</div>
          </div>
          <div
            className={styles.leftBody}
            ref={leftBodyRef}
            onScroll={handleLeftScroll}
            style={{ overflowY: 'auto' }}
          >
            {renderLeftRows()}
            <button className={styles.addTaskRowBtn} onClick={onAddTask}>
              {strings.Gantt_AddTaskButton}
            </button>
          </div>
        </div>

        {/* Right timeline panel */}
        <div className={styles.rightPanel}>
          {/* Sticky header */}
          <div className={styles.timelineHeaderScroll} ref={headerScrollRef}>
            <svg width={svgWidth} height={HEADER_HEIGHT} style={{ display: 'block', background: theme.bg }}>
              {/* Month row */}
              {monthBands.map((band, i) => (
                <g key={`month-${i}`}>
                  <rect
                    x={band.x}
                    y={0}
                    width={band.width}
                    height={28}
                    fill="transparent"
                  />
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
              {zoomLevel === 'day' ? (
                dayTicks.map((tick, i) => (
                  <g key={`day-${i}`}>
                    {tick.isWeekend && (
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
              width={svgWidth}
              height={Math.max(svgBodyHeight + 60, 400)}
              style={{ display: 'block' }}
              onMouseLeave={() => setTooltip(null)}
            >
              <defs>
                <marker id={`${uid}-arrow`} markerWidth="4" markerHeight="3" refX="3.5" refY="1.5" orient="auto">
                  <polygon points="0 0, 4 1.5, 0 3" fill="#8A8886" />
                </marker>
                <marker id={`${uid}-arrow-crit`} markerWidth="4" markerHeight="3" refX="3.5" refY="1.5" orient="auto">
                  <polygon points="0 0, 4 1.5, 0 3" fill="#D13438" />
                </marker>
                {settings.barStyle !== 'flat' && distinctBarColors.map(c => (
                  <linearGradient key={c} id={`${uid}-grad-${colorId(c)}`} x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor={c} stopOpacity="1" />
                    <stop offset="100%" stopColor={c} stopOpacity="0.75" />
                  </linearGradient>
                ))}
              </defs>

              {/* Weekend columns */}
              {weekendCols.map((col, i) => (
                <rect
                  key={`wknd-${i}`}
                  x={col.x}
                  y={0}
                  width={col.width}
                  height={svgBodyHeight + 60}
                  fill="#F8F7F6"
                />
              ))}

              {/* Row backgrounds + horizontal lines */}
              {visibleTasks.map((row, i) => (
                <g key={`rowbg-${i}`}>
                  {row.type === 'phase' && (
                    <rect
                      x={0}
                      y={i * ROW_H}
                      width={svgWidth}
                      height={ROW_H}
                      fill="#F3F2F1"
                    />
                  )}
                  <line
                    x1={0}
                    y1={(i + 1) * ROW_H}
                    x2={svgWidth}
                    y2={(i + 1) * ROW_H}
                    stroke="#F3F2F1"
                    strokeWidth={1}
                  />
                </g>
              ))}

              {/* Today highlight column */}
              <rect
                x={dateToX(today)}
                y={0}
                width={dayWidth}
                height={svgBodyHeight + 60}
                fill="rgba(255, 100, 100, 0.04)"
              />

              {/* Dependency arrows behind the bars */}
              {settings.showDependencies && renderDependencyArrows()}

              {/* Task bars */}
              {visibleTasks.map((row, i) => {
                if (row.type !== 'task' || !row.task) return null;
                return renderTaskBar(row.task, i);
              })}

              {/* Today line */}
              <line
                x1={todayX}
                y1={0}
                x2={todayX}
                y2={svgBodyHeight + 60}
                stroke="#D13438"
                strokeWidth={2}
                strokeDasharray="4,3"
                className={styles.todayLine}
              />
              <circle cx={todayX} cy={4} r={4} fill="#D13438" />
            </svg>
          </div>
        </div>
      </div>

      {/* Tooltip */}
      {tooltip && !dragState && (
        <div
          className={styles.tooltip}
          style={{
            left: tooltip.x + 14,
            top: tooltip.y - 10,
          }}
        >
          <div className={styles.tooltipTitle}>{tooltip.task.title}</div>
          <div className={styles.tooltipRow}>
            <span>📅</span>
            <span>
              {formatDateOnly(tooltip.task.startDate, 'MMM d, yyyy')} → {formatDateOnly(tooltip.task.dueDate, 'MMM d, yyyy')}
            </span>
          </div>
          <div className={styles.tooltipRow}>
            <span
              style={{
                display: 'inline-block',
                width: 8,
                height: 8,
                borderRadius: '50%',
                background: STATUS_COLORS[tooltip.task.status],
              }}
            />
            <span>{tooltip.task.status}</span>
            <span style={{ marginLeft: 8 }}>▪</span>
            <span>{tooltip.task.priority}</span>
          </div>
          {tooltip.task.assignedTo && (
            <div className={styles.tooltipRow}>
              <span>👤</span>
              <span>{tooltip.task.assignedTo}</span>
            </div>
          )}
          <div className={styles.tooltipRow}>
            <span>⬛</span>
            <span>{formatString(strings.Gantt_PercentCompleteTooltip, { value: tooltip.task.percentComplete })}</span>
          </div>
          {settings.showHealthBadges && (
            <div className={styles.tooltipRow}>
              <HealthBadge health={computeTaskHealth(tooltip.task)} size="md" />
            </div>
          )}
          {tooltip.task.phase && (
            <div className={styles.tooltipRow}>
              <span>🏷</span>
              <span>{tooltip.task.phase}</span>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

// ─── Helpers ─────────────────────────────────────────────────────────────────

interface IVisibleRow {
  type: 'task' | 'phase';
  task?: ITask;
  phase?: string;
  isChild?: boolean;
}

function buildVisibleRows(tasks: ITask[], collapsedPhases: Set<string>): IVisibleRow[] {
  const rows: IVisibleRow[] = [];
  const ids = new Set(tasks.map(t => t.id));
  // A task only renders as a sub-task if its parent is actually present;
  // otherwise (parent deleted or filtered out) it's promoted to top level so
  // it never silently disappears.
  const isSubTask = (t: ITask): boolean => !!t.parentTaskId && ids.has(t.parentTaskId);

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

  // Recurse through every nesting level (a sub-task can itself have
  // sub-tasks — e.g. via a direct list edit or import, even though the Task
  // panel's parent picker prevents creating it through the UI) so deeper
  // descendants are flattened into view instead of silently vanishing.
  // `visited` guards against a parent-cycle in the raw data looping forever.
  const addTask = (task: ITask, visited: Set<number> = new Set()): void => {
    rows.push({ type: 'task', task, isChild: visited.size > 0 });
    if (visited.has(task.id)) return;
    visited.add(task.id);
    const children = subtaskMap.get(task.id);
    if (children) {
      children.forEach(c => addTask(c, visited));
    }
  };

  byPhase.forEach((phaseTasks, phase) => {
    rows.push({ type: 'phase', phase });
    if (!collapsedPhases.has(phase)) {
      phaseTasks.forEach(t => addTask(t));
    }
  });

  noPhase.forEach(t => addTask(t));

  return rows;
}

export default GanttChart;
