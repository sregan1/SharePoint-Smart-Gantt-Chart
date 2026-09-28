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
import styles from './KanbanView.module.scss';

interface IKanbanViewProps {
  tasks: ITask[];
  project: IProject;
  showHealthBadges?: boolean;
  onEditTask: (task: ITask) => void;
  onDeleteTask: (id: number) => void;
  onTaskUpdate: (id: number, updates: Partial<ITask>) => void;
  onAddTask: () => void;
}

interface IColumn {
  status: TaskStatus;
  label: string;
}

// Colors come from the shared STATUS_COLORS map (models/index.ts) rather
// than being hardcoded here a second time.
const COLUMNS: IColumn[] = [
  { status: 'Not Started', label: strings.Kanban_ColumnNotStarted },
  { status: 'In Progress', label: strings.Kanban_ColumnInProgress },
  { status: 'On Hold', label: strings.Kanban_ColumnOnHold },
  { status: 'Completed', label: strings.Kanban_ColumnCompleted },
];

// Left/Right arrow keys step through this same order — drag-and-drop is
// otherwise the only way to move a card between columns, which a keyboard
// or screen-reader user can't do at all.
const STATUS_SEQUENCE: TaskStatus[] = [...COLUMNS.map(c => c.status), 'Cancelled'];

export const KanbanView: React.FC<IKanbanViewProps> = ({
  tasks, showHealthBadges = true, onEditTask, onDeleteTask, onTaskUpdate, onAddTask,
}) => {
  const [draggingId, setDraggingId] = React.useState<number | null>(null);
  const [dragOverCol, setDragOverCol] = React.useState<TaskStatus | null>(null);

  // Group tasks by status. Cancelled gets its own bucket — without it,
  // cancelled tasks used to fall through to Not Started AND render again in
  // the Cancelled column. Sub-tasks with a present parent stay off the board;
  // orphaned sub-tasks are shown so they never silently disappear.
  const tasksByStatus = React.useMemo(() => {
    const ids = new Set(tasks.map(t => t.id));
    const isSubTask = (t: ITask): boolean => !!t.parentTaskId && ids.has(t.parentTaskId);
    const map = new Map<TaskStatus, ITask[]>();
    COLUMNS.forEach(c => map.set(c.status, []));
    map.set('Cancelled', []);
    tasks
      .filter(t => !isSubTask(t))
      .forEach(t => {
        (map.get(t.status) || map.get('Not Started')!).push(t);
      });
    return map;
  }, [tasks]);

  const handleDragStart = (e: React.DragEvent, taskId: number): void => {
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('taskId', String(taskId));
    setDraggingId(taskId);
  };

  const handleDragEnd = (): void => {
    setDraggingId(null);
    setDragOverCol(null);
  };

  const handleDragOver = (e: React.DragEvent, status: TaskStatus): void => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    setDragOverCol(status);
  };

  const handleDragLeave = (e: React.DragEvent): void => {
    // dragleave fires when the pointer moves onto a descendant card, not
    // just when it truly leaves the column — clearing unconditionally made
    // the drop highlight/placeholder flicker while dragging across cards.
    if (e.currentTarget.contains(e.relatedTarget as Node)) return;
    setDragOverCol(null);
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
    setDragOverCol(null);
  };

  const handleCardKeyDown = (e: React.KeyboardEvent, task: ITask): void => {
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
            <span className={styles.cardTag} style={{ background: '#F3F2F1', color: '#605E5C' }}>
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
              <span style={{ fontSize: 10, color: '#605E5C', flexShrink: 0 }}>
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

  const cancelledTasks = tasksByStatus.get('Cancelled') || [];

  return (
    <div className={styles.kanbanView}>
      <div className={styles.board}>
        {COLUMNS.map(col => {
          const colTasks = tasksByStatus.get(col.status) || [];
          const isDragOver = dragOverCol === col.status;

          return (
            <div
              key={col.status}
              className={`${styles.column} ${isDragOver ? styles.dragOver : ''}`}
              onDragOver={e => handleDragOver(e, col.status)}
              onDragLeave={handleDragLeave}
              onDrop={e => handleDrop(e, col.status)}
            >
              {/* Column header */}
              <div className={styles.columnHeader}>
                <div className={styles.columnDot} style={{ background: STATUS_COLORS[col.status] }} />
                <span className={styles.columnTitle}>{col.label}</span>
                <span className={styles.columnCount}>{colTasks.length}</span>
              </div>

              {/* Cards */}
              <div className={styles.cardList}>
                {isDragOver && draggingId !== null && (
                  <div className={styles.dropPlaceholder} />
                )}
                {colTasks.map(task => renderCard(task, STATUS_COLORS[col.status]))}
              </div>

              {/* Add card button */}
              <button
                className={styles.addCardBtn}
                onClick={onAddTask}
              >
                {strings.Kanban_AddTaskButton}
              </button>
            </div>
          );
        })}

        {/* Cancelled column — collapsed sidebar style */}
        <div
          className={`${styles.column} ${dragOverCol === 'Cancelled' ? styles.dragOver : ''}`}
          style={{ opacity: 0.6 }}
          onDragOver={e => handleDragOver(e, 'Cancelled')}
          onDragLeave={handleDragLeave}
          onDrop={e => handleDrop(e, 'Cancelled')}
        >
          <div className={styles.columnHeader}>
            <div className={styles.columnDot} style={{ background: STATUS_COLORS['Cancelled'] }} />
            <span className={styles.columnTitle}>{strings.Kanban_ColumnCancelled}</span>
            <span className={styles.columnCount}>{cancelledTasks.length}</span>
          </div>
          <div className={styles.cardList}>
            {dragOverCol === 'Cancelled' && draggingId !== null && (
              <div className={styles.dropPlaceholder} />
            )}
            {cancelledTasks.map(task => renderCard(task, STATUS_COLORS['Cancelled']))}
          </div>
        </div>
      </div>
    </div>
  );
};

export default KanbanView;
