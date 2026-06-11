import * as React from 'react';
import {
  ITask, IProject, TaskStatus,
  STATUS_COLORS, STATUS_LIGHT_COLORS, PRIORITY_COLORS,
} from '../../models';
import { computeTaskHealth } from '../../utils/healthUtils';
import { parseDateOnly, formatDateOnly, todayLocalMidnight } from '../../utils/dateUtils';
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
  color: string;
}

const COLUMNS: IColumn[] = [
  { status: 'Not Started', label: 'Not Started', color: '#8B929A' },
  { status: 'In Progress', label: 'In Progress', color: '#0078D4' },
  { status: 'On Hold', label: 'On Hold', color: '#CA5010' },
  { status: 'Completed', label: 'Completed', color: '#107C10' },
];

function isOverdue(task: ITask): boolean {
  if (!task.dueDate || task.status === 'Completed' || task.status === 'Cancelled') return false;
  const due = parseDateOnly(task.dueDate);
  return !!due && due < todayLocalMidnight();
}

function initials(name: string): string {
  if (!name) return '?';
  return name.split(' ').map(p => p[0]).join('').substring(0, 2).toUpperCase();
}

function stringToColor(s: string): string {
  let hash = 0;
  for (let i = 0; i < s.length; i++) hash = s.charCodeAt(i) + ((hash << 5) - hash);
  return `hsl(${Math.abs(hash) % 360}, 55%, 45%)`;
}

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

  const handleDragLeave = (): void => {
    setDragOverCol(null);
  };

  const handleDrop = (e: React.DragEvent, newStatus: TaskStatus): void => {
    e.preventDefault();
    const taskId = parseInt(e.dataTransfer.getData('taskId'), 10);
    if (!isNaN(taskId)) {
      const task = tasks.find(t => t.id === taskId);
      if (task && task.status !== newStatus) {
        const updates: Partial<ITask> = { status: newStatus };
        if (newStatus === 'Completed' && task.percentComplete < 100) {
          updates.percentComplete = 100;
        }
        if (newStatus === 'Not Started' && task.percentComplete > 0) {
          updates.percentComplete = 0;
        }
        onTaskUpdate(taskId, updates);
      }
    }
    setDraggingId(null);
    setDragOverCol(null);
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
      >
        {/* Quick actions */}
        <div className={styles.cardActions}>
          <button
            className={styles.cardActionBtn}
            onClick={e => { e.stopPropagation(); onEditTask(task); }}
            title="Edit"
            aria-label={`Edit task ${task.title}`}
          >
            ✏
          </button>
          <button
            className={`${styles.cardActionBtn} ${styles.deleteBtn}`}
            onClick={e => { e.stopPropagation(); onDeleteTask(task.id); }}
            title="Delete"
            aria-label={`Delete task ${task.title}`}
          >
            ✕
          </button>
        </div>

        {/* Card header */}
        <div className={styles.cardHeader}>
          <div
            className={styles.priorityDot}
            style={{ background: PRIORITY_COLORS[task.priority] }}
            title={task.priority}
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
            {task.status}
          </span>
          <span
            className={styles.cardTag}
            style={{
              background: `${PRIORITY_COLORS[task.priority]}18`,
              color: PRIORITY_COLORS[task.priority],
              border: `1px solid ${PRIORITY_COLORS[task.priority]}40`,
            }}
          >
            {task.priority}
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
                <div className={styles.columnDot} style={{ background: col.color }} />
                <span className={styles.columnTitle}>{col.label}</span>
                <span className={styles.columnCount}>{colTasks.length}</span>
              </div>

              {/* Cards */}
              <div className={styles.cardList}>
                {isDragOver && draggingId !== null && (
                  <div className={styles.dropPlaceholder} />
                )}
                {colTasks.map(task => renderCard(task, col.color))}
              </div>

              {/* Add card button */}
              <button
                className={styles.addCardBtn}
                onClick={onAddTask}
              >
                + Add Task
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
            <span className={styles.columnTitle}>Cancelled</span>
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
