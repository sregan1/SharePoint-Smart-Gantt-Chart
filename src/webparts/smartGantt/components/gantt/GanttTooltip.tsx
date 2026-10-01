import * as React from 'react';
import { differenceInCalendarDays } from 'date-fns';
import { ITask, STATUS_COLORS } from '../../models';
import { computeTaskHealth } from '../../utils/healthUtils';
import { formatDateOnly, parseDateOnly } from '../../utils/dateUtils';
import { HealthBadge } from '../common/HealthBadge';
import * as strings from 'SmartGanttWebPartStrings';
import { formatString } from '../localeUtils';
import styles from './GanttChart.module.scss';

/** Imperative handle so hovering a bar updates only this small component. */
export interface ITooltipApi {
  show: (task: ITask, x: number, y: number) => void;
  hide: () => void;
}

interface IGanttTooltipProps {
  apiRef: React.MutableRefObject<ITooltipApi | null>;
  showHealthBadges: boolean;
}

interface ITooltipState { x: number; y: number; task: ITask }

// Owns the hovered-task state. Keeping it out of GanttChart means moving the
// pointer across bars re-renders this one node instead of every bar and row.
export const GanttTooltip: React.FC<IGanttTooltipProps> = ({ apiRef, showHealthBadges }) => {
  const [tip, setTip] = React.useState<ITooltipState | null>(null);

  React.useEffect(() => {
    apiRef.current = {
      show: (task, x, y) => setTip({ task, x, y }),
      hide: () => setTip(null),
    };
    return () => { apiRef.current = null; };
  }, [apiRef]);

  if (!tip) return null;
  const task = tip.task;

  const hasBaseline = !!task.baselineStart && !!task.baselineDue;
  let variance = 0;
  if (hasBaseline) {
    const due = parseDateOnly(task.dueDate);
    const baseDue = parseDateOnly(task.baselineDue);
    if (due && baseDue) variance = differenceInCalendarDays(due, baseDue);
  }

  // Keep the tooltip on screen when the bar sits near the right/bottom edge.
  const vw = typeof window !== 'undefined' ? window.innerWidth : 1200;
  const vh = typeof window !== 'undefined' ? window.innerHeight : 800;
  const left = Math.max(4, Math.min(tip.x + 14, vw - 260));
  const top = Math.max(4, Math.min(tip.y - 10, vh - 200));

  return (
    <div className={styles.tooltip} style={{ left, top }} role="tooltip">
      <div className={styles.tooltipTitle}>{task.title}</div>
      <div className={styles.tooltipRow}>
        <span>📅</span>
        <span>
          {formatDateOnly(task.startDate, 'MMM d, yyyy')} → {formatDateOnly(task.dueDate, 'MMM d, yyyy')}
        </span>
      </div>
      <div className={styles.tooltipRow}>
        <span
          style={{
            display: 'inline-block',
            width: 8,
            height: 8,
            borderRadius: '50%',
            background: STATUS_COLORS[task.status],
          }}
        />
        <span>{task.status}</span>
        <span style={{ marginLeft: 8 }}>▪</span>
        <span>{task.priority}</span>
      </div>
      {task.assignedTo && (
        <div className={styles.tooltipRow}>
          <span>👤</span>
          <span>{task.assignedTo}</span>
        </div>
      )}
      <div className={styles.tooltipRow}>
        <span>⬛</span>
        <span>{formatString(strings.Gantt_PercentCompleteTooltip, { value: task.percentComplete })}</span>
      </div>
      {hasBaseline && (
        <>
          <div className={styles.tooltipRow}>
            <span>📐</span>
            <span>
              {formatString(strings.Gantt_BaselineTooltip, {
                start: formatDateOnly(task.baselineStart, 'MMM d, yyyy'),
                due: formatDateOnly(task.baselineDue, 'MMM d, yyyy'),
              })}
            </span>
          </div>
          <div className={styles.tooltipRow}>
            <span>±</span>
            <span>
              {formatString(strings.Gantt_VarianceTooltip, { days: variance > 0 ? `+${variance}` : String(variance) })}
            </span>
          </div>
        </>
      )}
      {showHealthBadges && (
        <div className={styles.tooltipRow}>
          <HealthBadge health={computeTaskHealth(task)} size="md" />
        </div>
      )}
      {task.phase && (
        <div className={styles.tooltipRow}>
          <span>🏷</span>
          <span>{task.phase}</span>
        </div>
      )}
    </div>
  );
};
