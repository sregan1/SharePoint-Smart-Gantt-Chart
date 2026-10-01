import * as React from 'react';
import { ITask } from '../../models';
import { formatDateOnly, dateToDateOnlyString } from '../../utils/dateUtils';
import * as strings from 'SmartGanttWebPartStrings';
import { formatString } from '../localeUtils';
import {
  BAR_HEIGHT, MILESTONE_SIZE, MIN_BAR_WIDTH, CRITICAL_COLOR, hexToRgba, colorId,
} from './ganttUtils';
import styles from './GanttChart.module.scss';

export type BarDragMode = 'move' | 'resize-start' | 'resize-end';

/**
 * Stable callbacks shared by every bar. GanttChart builds this object once
 * (reading current props through a ref), which is what lets the memoized
 * bars skip re-rendering while another bar is dragged or hovered.
 */
export interface IBarHandlers {
  onPointerDown: (e: React.PointerEvent, task: ITask, mode: BarDragMode) => void;
  onLinkStart: (e: React.PointerEvent, task: ITask, fromX: number, fromY: number) => void;
  onActivate: (task: ITask) => void;
  onKeyDown: (e: React.KeyboardEvent, task: ITask) => void;
  onHover: (task: ITask, x: number, y: number) => void;
  onLeave: () => void;
}

interface IGanttBarProps {
  task: ITask;
  rowIndex: number;
  rowH: number;
  /** Left edge of a bar; horizontal center of a milestone diamond. */
  x: number;
  width: number;
  color: string;
  isMilestone: boolean;
  highlightCritical: boolean;
  flat: boolean;
  showProgressText: boolean;
  showAssignee: boolean;
  isLinkTarget: boolean;
  uid: string;
  hintId: string;
  /** Baseline geometry as primitives so a bar's memo isn't defeated by a fresh object. */
  baselineX?: number;
  baselineWidth?: number;
  /** Finish is later than the baseline finish. */
  baselineSlipped?: boolean;
  handlers: IBarHandlers;
}

const GanttBarComponent: React.FC<IGanttBarProps> = ({
  task, rowIndex, rowH, x, width, color, isMilestone, highlightCritical, flat,
  showProgressText, showAssignee, isLinkTarget, uid, hintId, baselineX, baselineWidth, baselineSlipped, handlers,
}) => {
  const { onPointerDown, onLinkStart, onActivate, onKeyDown, onHover, onLeave } = handlers;

  const showTip = (e: React.MouseEvent): void => onHover(task, e.clientX, e.clientY);
  // Keyboard users get the same tooltip as mouse users.
  const showTipOnFocus = (e: React.FocusEvent): void => {
    const r = (e.currentTarget as unknown as Element).getBoundingClientRect();
    onHover(task, r.left + r.width / 2, r.bottom);
  };

  const critStyle: React.CSSProperties | undefined = highlightCritical
    ? { stroke: CRITICAL_COLOR, strokeWidth: 1.5, strokeDasharray: '4,2' }
    : undefined;

  const baselineEl = baselineX !== undefined && (
    <rect
      x={baselineX}
      y={rowIndex * rowH + (rowH - BAR_HEIGHT) / 2 + BAR_HEIGHT + 1}
      width={Math.max(baselineWidth || 0, MIN_BAR_WIDTH)}
      height={3}
      rx={1.5}
      className={baselineSlipped ? styles.baselineBarSlipped : styles.baselineBar}
      pointerEvents="none"
    />
  );

  if (isMilestone) {
    const my = rowIndex * rowH + rowH / 2;
    const dateIso = task.startDate || task.dueDate;
    const ariaLabel = formatString(strings.Gantt_MilestoneAriaLabel, {
      taskTitle: task.title,
      date: formatDateOnly(dateIso, 'MMM d, yyyy'),
    });
    return (
      <>
        {baselineEl}
        <g
          className={`${styles.taskBarGroup} ${isLinkTarget ? styles.linkTarget : ''}`}
          role="button"
          tabIndex={0}
          aria-label={ariaLabel}
          aria-describedby={hintId}
          onClick={() => onActivate(task)}
          onKeyDown={e => onKeyDown(e, task)}
          onMouseEnter={showTip}
          onMouseLeave={onLeave}
          onFocus={showTipOnFocus}
          onBlur={onLeave}
        >
          <polygon
            className={styles.milestoneShape}
            points={`${x},${my - MILESTONE_SIZE} ${x + MILESTONE_SIZE},${my} ${x},${my + MILESTONE_SIZE} ${x - MILESTONE_SIZE},${my}`}
            fill={color}
            strokeWidth="1.5"
            style={critStyle ? { stroke: CRITICAL_COLOR, touchAction: 'none' } : { touchAction: 'none' }}
            onPointerDown={e => onPointerDown(e, task, 'move')}
          />
          <circle
            className={styles.taskConnector}
            cx={x + MILESTONE_SIZE + 8}
            cy={my}
            r={5}
            onPointerDown={e => onLinkStart(e, task, x + MILESTONE_SIZE + 8, my)}
          >
            <title>{strings.Gantt_LinkHandleTitle}</title>
          </circle>
        </g>
      </>
    );
  }

  const y = rowIndex * rowH + (rowH - BAR_HEIGHT) / 2;
  const barWidth = Math.max(MIN_BAR_WIDTH, width);
  const progressWidth = barWidth * (task.percentComplete / 100);
  // "Flat" bar style fills with the plain color; "Gradient" (default)
  // references the shared per-color gradient hoisted into the SVG's <defs>.
  const progressFill = flat ? color : `url(#${uid}-grad-${colorId(color)})`;
  const ariaLabel = formatString(strings.Gantt_BarAriaLabel, {
    taskTitle: task.title,
    startDate: formatDateOnly(task.startDate || dateToDateOnlyString(new Date()), 'MMM d, yyyy'),
    dueDate: formatDateOnly(task.dueDate, 'MMM d, yyyy'),
    percent: task.percentComplete,
  });
  const connectorX = x + barWidth + 8;
  const cy = y + BAR_HEIGHT / 2;

  return (
    <>
      {baselineEl}
      <g
        className={`${styles.taskBarGroup} ${isLinkTarget ? styles.linkTarget : ''}`}
        role="button"
        tabIndex={0}
        aria-label={ariaLabel}
        aria-describedby={hintId}
        onKeyDown={e => onKeyDown(e, task)}
        onFocus={showTipOnFocus}
        onBlur={onLeave}
      >
        {/* Background bar */}
        <rect
          x={x}
          y={y}
          width={barWidth}
          height={BAR_HEIGHT}
          rx={4}
          fill={hexToRgba(color, 0.18)}
          style={critStyle ? { ...critStyle, touchAction: 'none' } : { touchAction: 'none' }}
          className={styles.taskBar}
          onPointerDown={e => onPointerDown(e, task, 'move')}
          onClick={() => onActivate(task)}
          onMouseEnter={showTip}
          onMouseLeave={onLeave}
        />
        {/* Progress fill */}
        {progressWidth > 0 && (
          <rect
            className={styles.taskProgress}
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
        {showProgressText && barWidth > 50 && task.percentComplete > 0 && (
          <text
            className={styles.barText}
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
        {showAssignee && task.assignedTo && (
          <text
            className={styles.assigneeText}
            x={connectorX + 10}
            y={y + BAR_HEIGHT / 2 + 4}
            fontSize={11}
            fontFamily="'Segoe UI', sans-serif"
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
          onPointerDown={e => onPointerDown(e, task, 'resize-start')}
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
          onPointerDown={e => onPointerDown(e, task, 'resize-end')}
          style={{ cursor: 'ew-resize', touchAction: 'none' }}
        />
        {/* Connector: drag onto another bar to create a dependency */}
        <circle
          className={styles.taskConnector}
          cx={connectorX}
          cy={cy}
          r={5}
          onPointerDown={e => onLinkStart(e, task, connectorX, cy)}
        >
          <title>{strings.Gantt_LinkHandleTitle}</title>
        </circle>
      </g>
    </>
  );
};

export const GanttBar = React.memo(GanttBarComponent);
