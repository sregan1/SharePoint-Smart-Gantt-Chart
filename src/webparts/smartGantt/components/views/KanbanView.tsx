import * as React from 'react';
import * as strings from 'SmartGanttWebPartStrings';
import { formatString } from '../localeUtils';
import {
  ITask, IProject, TaskStatus,
  STATUS_COLORS, STATUS_LIGHT_COLORS, PRIORITY_COLORS,
} from '../../models';
import { computeTaskHealth } from '../../utils/healthUtils';
import { formatDateOnly } from '../../utils/dateUtils';
import { isOverdue, initials, stringToColor, getStatusLabel, getPriorityLabel } from '../../utils/taskDisplayUtils';
import { HealthBadge } from '../common/HealthBadge';
import { readStored, writeStored } from '../common/storage';
import styles from './KanbanView.module.scss';

interface IKanbanViewProps {
  tasks: ITask[];
  project: IProject;
  showHealthBadges?: boolean;
  onEditTask: (task: ITask) => void;
  onDeleteTask: (id: number) => void;
  onTaskUpdate: (id: number, updates: Partial<ITask>) => void;
  /** Called from a column's "+ Add Task" button with that column's status so
   *  the new task can be pre-set to it; called without arguments elsewhere. */
  onAddTask: (status?: TaskStatus) => void;
}

interface IColumn {
  status: TaskStatus;
  label: string;
}

type LaneMode = 'none' | 'assignee' | 'phase';
type WipLimits = Partial<Record<TaskStatus, number>>;

interface ILane {
  key: string;
  label: string;
  tasks: ITask[];
}

// Colors come from the shared STATUS_COLORS map (models/index.ts) rather
// than being hardcoded here a second time.
const COLUMNS: IColumn[] = [
  { status: 'Not Started', label: strings.Kanban_ColumnNotStarted },
  { status: 'In Progress', label: strings.Kanban_ColumnInProgress },
  { status: 'On Hold', label: strings.Kanban_ColumnOnHold },
  { status: 'Completed', label: strings.Kanban_ColumnCompleted },
];

const CANCELLED_COLUMN: IColumn = { status: 'Cancelled', label: strings.Kanban_ColumnCancelled };

// Left/Right arrow keys step through this same order — drag-and-drop is
// otherwise the only way to move a card between columns, which a keyboard
// or screen-reader user can't do at all.
const STATUS_SEQUENCE: TaskStatus[] = [...COLUMNS.map(c => c.status), 'Cancelled'];

// Per-project view preferences, remembered in localStorage (best-effort).
const wipKey = (projectId: number): string => `sg.kanban.wip.${projectId}`;
const cancelledKey = (projectId: number): string => `sg.kanban.cancelledCollapsed.${projectId}`;
const laneKey = (projectId: number): string => `sg.kanban.lanes.${projectId}`;

export const KanbanView: React.FC<IKanbanViewProps> = ({
  tasks, project, showHealthBadges = true, onEditTask, onDeleteTask, onTaskUpdate, onAddTask,
}) => {
  const [draggingId, setDraggingId] = React.useState<number | null>(null);
  const [dragOver, setDragOver] = React.useState<string | null>(null);

  const [wipLimits, setWipLimits] = React.useState<WipLimits>(() => readStored<WipLimits>(wipKey(project.id), {}));
  const [cancelledCollapsed, setCancelledCollapsed] = React.useState<boolean>(() => readStored<boolean>(cancelledKey(project.id), false));
  const [laneMode, setLaneMode] = React.useState<LaneMode>(() => readStored<LaneMode>(laneKey(project.id), 'none'));
  const [editingWip, setEditingWip] = React.useState<string | null>(null);

  // Re-read the remembered preferences when the selected project changes.
  React.useEffect(() => {
    setWipLimits(readStored<WipLimits>(wipKey(project.id), {}));
    setCancelledCollapsed(readStored<boolean>(cancelledKey(project.id), false));
    setLaneMode(readStored<LaneMode>(laneKey(project.id), 'none'));
    setEditingWip(null);
  }, [project.id]);

  const commitWip = (status: TaskStatus, raw: string): void => {
    const n = parseInt(raw, 10);
    const next: WipLimits = { ...wipLimits };
    if (isNaN(n) || n <= 0) delete next[status]; else next[status] = n;
    setWipLimits(next);
    writeStored(wipKey(project.id), next);
    setEditingWip(null);
  };

  const toggleCancelled = (): void => {
    const next = !cancelledCollapsed;
    setCancelledCollapsed(next);
    writeStored(cancelledKey(project.id), next);
  };

  const changeLaneMode = (mode: LaneMode): void => {
    setLaneMode(mode);
    writeStored(laneKey(project.id), mode);
  };

  // Tasks that appear on the board. Sub-tasks with a present parent stay off
  // the board; orphaned sub-tasks are shown so they never silently disappear.
  const boardTasks = React.useMemo(() => {
    const ids = new Set(tasks.map(t => t.id));
    return tasks.filter(t => !(!!t.parentTaskId && ids.has(t.parentTaskId)));
  }, [tasks]);

  // Group tasks by status. Cancelled gets its own bucket — without it,
  // cancelled tasks used to fall through to Not Started AND render again in
  // the Cancelled column.
  const groupByStatus = (list: ITask[]): Map<TaskStatus, ITask[]> => {
    const map = new Map<TaskStatus, ITask[]>();
    COLUMNS.forEach(c => map.set(c.status, []));
    map.set('Cancelled', []);
    list.forEach(t => {
      (map.get(t.status) || map.get('Not Started')!).push(t);
    });
    return map;
  };

  const totalsByStatus = React.useMemo(() => groupByStatus(boardTasks), [boardTasks]);

  // Swimlanes: rows of the same status columns grouped by assignee or phase.
  // Named lanes sort alphabetically; the blank bucket always goes last.
  const lanes: ILane[] = React.useMemo(() => {
    if (laneMode === 'none') return [{ key: 'all', label: '', tasks: boardTasks }];
    const groups = new Map<string, ITask[]>();
    boardTasks.forEach(t => {
      const k = (laneMode === 'assignee' ? t.assignedTo : t.phase) || '';
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k)!.push(t);
    });
    const names = Array.from(groups.keys()).filter(k => k !== '')
      .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
    const blankLabel = laneMode === 'assignee' ? strings.View_Kanban_LaneUnassigned : strings.View_Kanban_LaneNoPhase;
    const result: ILane[] = names.map(n => ({ key: `n:${n}`, label: n, tasks: groups.get(n)! }));
    if (groups.has('')) result.push({ key: 'blank', label: blankLabel, tasks: groups.get('')! });
    return result;
  }, [boardTasks, laneMode]);

  const handleDragStart = (e: React.DragEvent, taskId: number): void => {
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('taskId', String(taskId));
    setDraggingId(taskId);
  };

  const handleDragEnd = (): void => {
    setDraggingId(null);
    setDragOver(null);
  };

  const handleDragOver = (e: React.DragEvent, key: string): void => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    setDragOver(key);
  };

  const handleDragLeave = (e: React.DragEvent): void => {
    // dragleave fires when the pointer moves onto a descendant card, not
    // just when it truly leaves the column — clearing unconditionally made
    // the drop highlight/placeholder flicker while dragging across cards.
    if (e.currentTarget.contains(e.relatedTarget as Node)) return;
    setDragOver(null);
  };

  // Shared by drag-drop and the keyboard path so both keep % complete
  // consistent with the new status the same way.
  const moveTaskToStatus = (task: ITask, newStatus: TaskStatus): void => {
    if (task.status === newStatus) return;
    const updates: Partial<ITask> = { status: newStatus };
    if (newStatus === 'Completed' && task.percentComplete < 100) {
      updates.percentComplete = 100;
    }
    if (newStatus === 'Not Started' && task.percentComplete > 0) {
      updates.percentComplete = 0;
    }
    onTaskUpdate(task.id, updates);
  };

  const handleDrop = (e: React.DragEvent, newStatus: TaskStatus): void => {
    e.preventDefault();
    const taskId = parseInt(e.dataTransfer.getData('taskId'), 10);
    if (!isNaN(taskId)) {
      const task = tasks.find(t => t.id === taskId);
      if (task) moveTaskToStatus(task, newStatus);
    }
    setDraggingId(null);
    setDragOver(null);
  };

  const handleCardKeyDown = (e: React.KeyboardEvent, task: ITask): void => {
    // Key presses on the card's inner Edit/Delete buttons bubble up here —
    // only react to keys pressed on the card itself.
    if (e.target !== e.currentTarget) return;
    if (e.key === 'Enter') { onEditTask(task); return; }
    if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
    const idx = STATUS_SEQUENCE.indexOf(task.status);
    if (idx === -1) return;
    const nextIdx = e.key === 'ArrowRight' ? idx + 1 : idx - 1;
    if (nextIdx < 0 || nextIdx >= STATUS_SEQUENCE.length) return;
    e.preventDefault();
    moveTaskToStatus(task, STATUS_SEQUENCE[nextIdx]);
  };

  const renderCard = (task: ITask, colColor: string): React.ReactNode => {
    const overdue = isOverdue(task);
    const dueStr = formatDateOnly(task.dueDate, 'MMM d', '');
    const startStr = formatDateOnly(task.startDate, 'MMM d', '');

    return (
      <div
        key={task.id}
        className={`${styles.card} ${task.isMilestone ? styles.milestone : ''} ${draggingId === task.id ? styles.dragging : ''}`}
        style={{ borderLeftColor: task.color || colColor }}
        draggable
        onDragStart={e => handleDragStart(e, task.id)}
        onDragEnd={handleDragEnd}
        tabIndex={0}
        role="button"
        aria-label={formatString(strings.Kanban_CardAriaLabel, { taskTitle: task.title, status: getStatusLabel(task.status) })}
        onKeyDown={e => handleCardKeyDown(e, task)}
      >
        {/* Quick actions */}
        <div className={styles.cardActions}>
          <button
            className={styles.cardActionBtn}
            onClick={e => { e.stopPropagation(); onEditTask(task); }}
            title={strings.Kanban_EditTitle}
            aria-label={formatString(strings.Kanban_EditTaskAriaLabel, { taskTitle: task.title })}
          >
            ✏
          </button>
          <button
            className={`${styles.cardActionBtn} ${styles.deleteBtn}`}
            onClick={e => { e.stopPropagation(); onDeleteTask(task.id); }}
            title={strings.Kanban_DeleteTitle}
            aria-label={formatString(strings.Kanban_DeleteTaskAriaLabel, { taskTitle: task.title })}
          >
            ✕
          </button>
        </div>

        {/* Card header */}
        <div className={styles.cardHeader}>
          <div
            className={styles.priorityDot}
            style={{ background: PRIORITY_COLORS[task.priority] }}
            title={getPriorityLabel(task.priority)}
          />
          {task.isMilestone && <span className={styles.cardMilestoneIcon}>◆</span>}
          <span className={styles.cardTitle} onClick={() => onEditTask(task)}>
            {task.title}
          </span>
        </div>

        {/* Tags */}
        <div className={styles.cardMeta}>
          <span
            className={styles.cardTag}
            style={{
              background: STATUS_LIGHT_COLORS[task.status],
              color: STATUS_COLORS[task.status],
            }}
          >
            {getStatusLabel(task.status)}
          </span>
          <span
            className={styles.cardTag}
            style={{
              background: `${PRIORITY_COLORS[task.priority]}18`,
              color: PRIORITY_COLORS[task.priority],
              border: `1px solid ${PRIORITY_COLORS[task.priority]}40`,
            }}
          >
            {getPriorityLabel(task.priority)}
          </span>
          {task.phase && (
            <span className={styles.cardTag} style={{ background: 'var(--neutralLighter, #F3F2F1)', color: 'var(--neutralSecondary, #605E5C)' }}>
              {task.phase}
            </span>
          )}
          {showHealthBadges && (
            <HealthBadge health={computeTaskHealth(task)} size="sm" />
          )}
        </div>

        {/* Dates */}
        {(startStr || dueStr) && (
          <div className={styles.cardMeta}>
            {startStr && (
              <span className={styles.cardDate}>
                📅 {startStr}
              </span>
            )}
            {dueStr && (
              <span className={`${styles.cardDate} ${overdue ? styles.overdue : ''}`}>
                {startStr ? '→' : '📅'} {dueStr}
                {overdue && ' ⚠'}
              </span>
            )}
          </div>
        )}

        {/* Footer: progress + avatar */}
        <div className={styles.cardFooter}>
          {task.percentComplete > 0 && (
            <>
              <div className={styles.progressBarSmall}>
                <div
                  className={styles.progressFill}
                  style={{
                    width: `${task.percentComplete}%`,
                    background: STATUS_COLORS[task.status],
                  }}
                />
              </div>
              <span style={{ fontSize: 10, color: 'var(--neutralSecondary, #605E5C)', flexShrink: 0 }}>
                {task.percentComplete}%
              </span>
            </>
          )}
          {task.assignedTo && (
            <div
              className={styles.assigneeAvatar}
              style={{ background: stringToColor(task.assignedTo) }}
              title={task.assignedTo}
            >
              {initials(task.assignedTo)}
            </div>
          )}
        </div>
      </div>
    );
  };

  // WIP control shown in each column header: a small button showing the
  // limit, swapped for a numeric input while editing. The limit is compared
  // against the column's total across all swimlanes.
  const renderWip = (col: IColumn, lane: string): React.ReactNode => {
    const limit = wipLimits[col.status];
    const total = (totalsByStatus.get(col.status) || []).length;
    const id = `${lane}|${col.status}`;
    if (editingWip === id) {
      return (
        <input
          type="number"
          min={0}
          className={styles.wipInput}
          defaultValue={limit || ''}
          autoFocus
          aria-label={formatString(strings.View_Kanban_WipInputAriaLabel, { column: col.label })}
          onBlur={e => commitWip(col.status, e.currentTarget.value)}
          onKeyDown={e => {
            if (e.key === 'Enter') { e.preventDefault(); commitWip(col.status, e.currentTarget.value); }
            else if (e.key === 'Escape') { e.stopPropagation(); setEditingWip(null); }
          }}
        />
      );
    }
    const title = formatString(strings.View_Kanban_WipEditTitle, { column: col.label });
    return (
      <button
        type="button"
        className={styles.wipBtn}
        onClick={() => setEditingWip(id)}
        title={title}
        aria-label={title}
      >
        {limit ? formatString(strings.View_Kanban_WipLabel, { count: total, limit }) : strings.View_Kanban_WipSet}
      </button>
    );
  };

  const renderColumn = (
    col: IColumn, lane: ILane, byStatus: Map<TaskStatus, ITask[]>, showAdd: boolean,
  ): React.ReactNode => {
    const colTasks = byStatus.get(col.status) || [];
    const dragKey = `${lane.key}|${col.status}`;
    const isDragOver = dragOver === dragKey;
    const limit = wipLimits[col.status];
    const total = (totalsByStatus.get(col.status) || []).length;
    const exceeded = !!limit && total > limit;

    return (
      <div
        key={col.status}
        className={`${styles.column} ${isDragOver ? styles.dragOver : ''}`}
        onDragOver={e => handleDragOver(e, dragKey)}
        onDragLeave={handleDragLeave}
        onDrop={e => handleDrop(e, col.status)}
      >
        {/* Column header */}
        <div
          className={`${styles.columnHeader} ${exceeded ? styles.wipExceeded : ''}`}
          title={exceeded ? formatString(strings.View_Kanban_WipExceededTitle, { count: total, limit: limit! }) : undefined}
        >
          <div className={styles.columnDot} style={{ background: STATUS_COLORS[col.status] }} />
          <span className={styles.columnTitle}>{col.label}</span>
          {exceeded && (
            <span aria-label={formatString(strings.View_Kanban_WipExceededTitle, { count: total, limit: limit! })} role="img">⚠</span>
          )}
          {renderWip(col, lane.key)}
          <span className={styles.columnCount}>{colTasks.length}</span>
        </div>

        {/* Cards */}
        <div className={styles.cardList}>
          {isDragOver && draggingId !== null && (
            <div className={styles.dropPlaceholder} />
          )}
          {colTasks.map(task => renderCard(task, STATUS_COLORS[col.status]))}
        </div>

        {/* Add card button — presets the new task's status to this column */}
        {showAdd && (
          <button
            className={styles.addCardBtn}
            onClick={() => onAddTask(col.status)}
            aria-label={formatString(strings.View_Kanban_AddTaskInColumnAriaLabel, { column: col.label })}
          >
            {strings.Kanban_AddTaskButton}
          </button>
        )}
      </div>
    );
  };

  const renderCancelled = (lane: ILane, byStatus: Map<TaskStatus, ITask[]>): React.ReactNode => {
    const cancelledTasks = byStatus.get('Cancelled') || [];
    const dragKey = `${lane.key}|Cancelled`;
    const isDragOver = dragOver === dragKey;
    const toggleLabel = formatString(
      cancelledCollapsed ? strings.View_Kanban_ExpandColumn : strings.View_Kanban_CollapseColumn,
      { column: CANCELLED_COLUMN.label },
    );
    const dropProps = {
      onDragOver: (e: React.DragEvent) => handleDragOver(e, dragKey),
      onDragLeave: handleDragLeave,
      onDrop: (e: React.DragEvent) => handleDrop(e, 'Cancelled'),
    };

    if (cancelledCollapsed) {
      // Still a drop target while collapsed so cards can be canceled by drag.
      return (
        <div
          key="Cancelled"
          className={`${styles.column} ${styles.columnCollapsed} ${isDragOver ? styles.dragOver : ''}`}
          {...dropProps}
        >
          <button type="button" className={styles.headerIconBtn} onClick={toggleCancelled} title={toggleLabel} aria-label={toggleLabel} aria-expanded={false}>
            ▸
          </button>
          <div className={styles.columnDot} style={{ background: STATUS_COLORS['Cancelled'] }} />
          <span className={styles.columnCount}>{cancelledTasks.length}</span>
          <span className={styles.collapsedTitle}>{CANCELLED_COLUMN.label}</span>
        </div>
      );
    }

    return (
      <div
        key="Cancelled"
        className={`${styles.column} ${isDragOver ? styles.dragOver : ''}`}
        style={{ opacity: 0.6 }}
        {...dropProps}
      >
        <div className={styles.columnHeader}>
          <div className={styles.columnDot} style={{ background: STATUS_COLORS['Cancelled'] }} />
          <span className={styles.columnTitle}>{CANCELLED_COLUMN.label}</span>
          <span className={styles.columnCount}>{cancelledTasks.length}</span>
          <button type="button" className={styles.headerIconBtn} onClick={toggleCancelled} title={toggleLabel} aria-label={toggleLabel} aria-expanded={true}>
            ◂
          </button>
        </div>
        <div className={styles.cardList}>
          {isDragOver && draggingId !== null && (
            <div className={styles.dropPlaceholder} />
          )}
          {cancelledTasks.map(task => renderCard(task, STATUS_COLORS['Cancelled']))}
        </div>
      </div>
    );
  };

  const renderBoard = (lane: ILane): React.ReactNode => {
    const byStatus = groupByStatus(lane.tasks);
    return (
      <div className={styles.board}>
        {COLUMNS.map(col => renderColumn(col, lane, byStatus, true))}
        {renderCancelled(lane, byStatus)}
      </div>
    );
  };

  const laneOptions: Array<{ id: LaneMode; label: string }> = [
    { id: 'none', label: strings.View_Kanban_LanesNone },
    { id: 'assignee', label: strings.View_Kanban_LanesAssignee },
    { id: 'phase', label: strings.View_Kanban_LanesPhase },
  ];

  return (
    <div className={styles.kanbanView}>
      <div className={styles.toolbar}>
        <label htmlFor={`kanban-lanes-${project.id}`}>{strings.View_Kanban_SwimlanesLabel}</label>
        <select
          id={`kanban-lanes-${project.id}`}
          value={laneMode}
          onChange={e => changeLaneMode(e.target.value as LaneMode)}
        >
          {laneOptions.map(o => <option key={o.id} value={o.id}>{o.label}</option>)}
        </select>
      </div>

      {laneMode === 'none' ? renderBoard(lanes[0]) : (
        <div className={styles.lanes}>
          {lanes.map(lane => (
            <div key={lane.key} className={styles.lane} role="group" aria-label={lane.label}>
              <div className={styles.laneHeader}>
                {formatString(strings.View_Kanban_LaneHeader, { name: lane.label, count: lane.tasks.length })}
              </div>
              {renderBoard(lane)}
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default KanbanView;
