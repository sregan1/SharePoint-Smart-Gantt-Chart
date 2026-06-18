import * as React from 'react';
import {
  ITask, IProject, TaskStatus, TaskPriority,
  STATUS_COLORS, STATUS_LIGHT_COLORS, PRIORITY_COLORS,
  TASK_STATUS_OPTIONS, TASK_PRIORITY_OPTIONS,
} from '../../models';
import { computeTaskHealth, hasDependencyViolation } from '../../utils/healthUtils';
import { parseDateOnly, formatDateOnly, todayLocalMidnight } from '../../utils/dateUtils';
import { HealthBadge } from '../common/HealthBadge';
import styles from './ListView.module.scss';

interface IListViewProps {
  tasks: ITask[];
  project: IProject;
  showHealthBadges?: boolean;
  onEditTask: (task: ITask) => void;
  onDeleteTask: (id: number) => void;
  onTaskUpdate: (id: number, updates: Partial<ITask>) => void;
  onAddTask: () => void;
}

type SortField = 'sortOrder' | 'title' | 'startDate' | 'dueDate' | 'status' | 'priority' | 'assignedTo' | 'percentComplete' | 'phase';
type SortDir = 'asc' | 'desc';

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
  const h = hash % 360;
  return `hsl(${Math.abs(h)}, 55%, 45%)`;
}

export const ListView: React.FC<IListViewProps> = ({
  tasks, showHealthBadges = true, onEditTask, onDeleteTask, onTaskUpdate, onAddTask,
}) => {
  const [sortField, setSortField] = React.useState<SortField>('sortOrder');
  const [sortDir, setSortDir] = React.useState<SortDir>('asc');

  const handleSort = (field: SortField): void => {
    if (sortField === field) {
      // asc → desc → back to manual order
      if (sortDir === 'asc') {
        setSortDir('desc');
      } else {
        setSortField('sortOrder');
        setSortDir('asc');
      }
    } else {
      setSortField(field);
      setSortDir('asc');
    }
  };

  const taskById = React.useMemo(
    () => new Map(tasks.map(t => [t.id, t])),
    [tasks]
  );

  const violationIds = React.useMemo(() => {
    const set = new Set<number>();
    tasks.forEach(t => { if (hasDependencyViolation(t, tasks)) set.add(t.id); });
    return set;
  }, [tasks]);

  const sortedTasks = React.useMemo(() => {
    const ids = new Set(tasks.map(t => t.id));
    // Only nest under a parent that's actually present; orphaned/filtered-out
    // parents must not make their sub-tasks disappear.
    const isSubTask = (t: ITask): boolean => !!t.parentTaskId && ids.has(t.parentTaskId);
    const top = tasks.filter(t => !isSubTask(t));
    const children = new Map<number, ITask[]>();
    tasks.filter(isSubTask).forEach(t => {
      if (!children.has(t.parentTaskId!)) children.set(t.parentTaskId!, []);
      children.get(t.parentTaskId!)!.push(t);
    });

    const sortFn = (a: ITask, b: ITask): number => {
      let cmp: number;
      if (sortField === 'startDate' || sortField === 'dueDate') {
        const av = parseDateOnly(a[sortField])?.getTime() ?? 0;
        const bv = parseDateOnly(b[sortField])?.getTime() ?? 0;
        cmp = av - bv;
      } else if (sortField === 'percentComplete' || sortField === 'sortOrder') {
        cmp = a[sortField] - b[sortField];
      } else {
        cmp = String(a[sortField] ?? '').localeCompare(String(b[sortField] ?? ''), undefined, { sensitivity: 'base' });
      }
      if (cmp === 0) cmp = a.id - b.id;
      return sortDir === 'asc' ? cmp : -cmp;
    };

    // Group by phase
    const byPhase = new Map<string, ITask[]>();
    const noPhase: ITask[] = [];
    top.forEach(t => {
      if (t.phase) {
        if (!byPhase.has(t.phase)) byPhase.set(t.phase, []);
        byPhase.get(t.phase)!.push(t);
      } else {
        noPhase.push(t);
      }
    });

    const rows: Array<{ type: 'task' | 'phase'; task?: ITask; phase?: string; isChild?: boolean }> = [];

    const pushTask = (t: ITask): void => {
      rows.push({ type: 'task', task: t, isChild: false });
      (children.get(t.id) || []).sort(sortFn).forEach(c => rows.push({ type: 'task', task: c, isChild: true }));
    };

    byPhase.forEach((pTasks, phase) => {
      rows.push({ type: 'phase', phase });
      [...pTasks].sort(sortFn).forEach(pushTask);
    });
    [...noPhase].sort(sortFn).forEach(pushTask);

    return rows;
  }, [tasks, sortField, sortDir]);

  const SortTh: React.FC<{ field: SortField; label: string; width?: number }> = ({ field, label, width }) => (
    <th
      className={sortField === field ? styles.sorted : ''}
      onClick={() => handleSort(field)}
      style={width ? { width } : undefined}
      role="columnheader"
      aria-sort={sortField === field ? (sortDir === 'asc' ? 'ascending' : 'descending') : 'none'}
    >
      {label}
      {sortField === field && (
        <span className={styles.sortArrow}>{sortDir === 'asc' ? '↑' : '↓'}</span>
      )}
    </th>
  );

  if (tasks.length === 0) {
    return (
      <div className={styles.listView}>
        <div className={styles.emptyState}>
          <div style={{ fontSize: 40, opacity: 0.3 }}>📋</div>
          <div style={{ fontSize: 16, fontWeight: 600, color: '#323130' }}>No tasks yet</div>
          <button
            style={{
              background: '#0078D4', color: '#fff', border: 'none', borderRadius: 4,
              padding: '8px 16px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
            }}
            onClick={onAddTask}
          >
            + Add First Task
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.listView}>
      <div className={styles.tableWrapper}>
        <table>
          <thead className={styles.thead}>
            <tr>
              <SortTh field="title" label="Task Name" />
              <SortTh field="status" label="Status" width={130} />
              {showHealthBadges && <th style={{ width: 100 }}>Health</th>}
              <SortTh field="priority" label="Priority" width={100} />
              <SortTh field="startDate" label="Start" width={110} />
              <SortTh field="dueDate" label="Due" width={110} />
              <SortTh field="assignedTo" label="Assigned To" width={140} />
              <SortTh field="percentComplete" label="Progress" width={140} />
              <SortTh field="phase" label="Phase" width={110} />
              <th style={{ width: 160 }}>Predecessors</th>
              <th style={{ width: 72 }} />
            </tr>
          </thead>
          <tbody className={styles.tbody}>
            {sortedTasks.map((row, _i) => {
              if (row.type === 'phase') {
                return (
                  <tr key={`phase-${row.phase}`} className={styles.phaseGroupRow}>
                    <td colSpan={showHealthBadges ? 11 : 10}>
                      <span className={styles.phaseGroupCell}>▸ {row.phase}</span>
                    </td>
                  </tr>
                );
              }

              const task = row.task!;
              const isChild = !!row.isChild;
              const overdue = isOverdue(task);

              return (
                <tr key={`task-${task.id}`}>
                  {/* Task name */}
                  <td>
                    <div className={styles.taskNameCell}>
                      {isChild && <div className={styles.subtaskIndent} />}
                      <div
                        className={styles.statusDot}
                        style={{ background: STATUS_COLORS[task.status] }}
                      />
                      {task.isMilestone && <span className={styles.milestoneIcon}>◆</span>}
                      <span
                        className={styles.taskName}
                        title={task.title}
                        onClick={() => onEditTask(task)}
                        style={{ cursor: 'pointer' }}
                      >
                        {task.title}
                      </span>
                      {violationIds.has(task.id) && (
                        <span
                          title="Started before all dependencies were completed"
                          style={{ color: '#CA5010', fontSize: 12, flexShrink: 0 }}
                        >
                          ⚠
                        </span>
                      )}
                    </div>
                  </td>

                  {/* Status */}
                  <td>
                    <select
                      className={styles.inlineSelect}
                      value={task.status}
                      aria-label={`Status of ${task.title}`}
                      onChange={e => {
                        const status = e.target.value as TaskStatus;
                        const updates: Partial<ITask> = { status };
                        // Keep % complete consistent with status, matching the
                        // behavior of the task panel and Kanban drop.
                        if (status === 'Completed' && task.percentComplete < 100) updates.percentComplete = 100;
                        if (status === 'Not Started' && task.percentComplete > 0) updates.percentComplete = 0;
                        onTaskUpdate(task.id, updates);
                      }}
                      style={{
                        background: STATUS_LIGHT_COLORS[task.status],
                        color: STATUS_COLORS[task.status],
                        fontWeight: 600,
                        borderRadius: 10,
                        paddingLeft: 8,
                        paddingRight: 8,
                      }}
                    >
                      {TASK_STATUS_OPTIONS.map(s => (
                        <option key={s} value={s}>{s}</option>
                      ))}
                    </select>
                  </td>

                  {/* Health */}
                  {showHealthBadges && (
                    <td>
                      <HealthBadge health={computeTaskHealth(task)} size="sm" />
                    </td>
                  )}

                  {/* Priority */}
                  <td>
                    <select
                      className={styles.inlineSelect}
                      value={task.priority}
                      aria-label={`Priority of ${task.title}`}
                      onChange={e => onTaskUpdate(task.id, { priority: e.target.value as TaskPriority })}
                      style={{
                        color: PRIORITY_COLORS[task.priority],
                        fontWeight: 600,
                      }}
                    >
                      {TASK_PRIORITY_OPTIONS.map(p => (
                        <option key={p} value={p}>{p}</option>
                      ))}
                    </select>
                  </td>

                  {/* Start */}
                  <td>
                    <span className={styles.dateCell}>{formatDateOnly(task.startDate, 'MMM d, yyyy')}</span>
                  </td>

                  {/* Due */}
                  <td>
                    <span className={`${styles.dateCell} ${overdue ? styles.overdue : ''}`}>
                      {formatDateOnly(task.dueDate, 'MMM d, yyyy')}
                      {overdue && ' ⚠'}
                    </span>
                  </td>

                  {/* Assigned to */}
                  <td>
                    {task.assignedTo ? (
                      <div className={styles.assigneeCell}>
                        <div
                          className={styles.avatarCircle}
                          style={{ background: stringToColor(task.assignedTo) }}
                          title={task.assignedTo}
                        >
                          {initials(task.assignedTo)}
                        </div>
                        <span style={{ fontSize: 12 }}>{task.assignedTo.split(' ')[0]}</span>
                      </div>
                    ) : (
                      <span style={{ color: '#C8C6C4', fontSize: 12 }}>Unassigned</span>
                    )}
                  </td>

                  {/* Progress */}
                  <td>
                    <div className={styles.progressCell}>
                      <div className={styles.progressBar}>
                        <div
                          className={styles.progressFill}
                          style={{
                            width: `${task.percentComplete}%`,
                            background: STATUS_COLORS[task.status],
                          }}
                        />
                      </div>
                      <span className={styles.progressLabel}>{task.percentComplete}%</span>
                    </div>
                  </td>

                  {/* Phase */}
                  <td>
                    <span style={{ fontSize: 12, color: '#605E5C' }}>
                      {task.phase || '—'}
                    </span>
                  </td>

                  {/* Predecessors */}
                  <td style={{ maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    <span style={{ fontSize: 12, color: '#605E5C' }} title={task.dependencies.map(id => taskById.get(id)?.title ?? `#${id}`).join(', ')}>
                      {task.dependencies.length > 0
                        ? task.dependencies.map(id => taskById.get(id)?.title ?? `#${id}`).join(', ')
                        : '—'}
                    </span>
                  </td>

                  {/* Actions */}
                  <td>
                    <div className={styles.rowActions}>
                      <button
                        className={styles.rowActionBtn}
                        onClick={() => onEditTask(task)}
                        title="Edit"
                        aria-label={`Edit task ${task.title}`}
                      >
                        ✏
                      </button>
                      <button
                        className={`${styles.rowActionBtn} ${styles.deleteBtn}`}
                        onClick={() => onDeleteTask(task.id)}
                        title="Delete"
                        aria-label={`Delete task ${task.title}`}
                      >
                        ✕
                      </button>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
        <button className={styles.addRowBtn} onClick={onAddTask}>
          + Add Task
        </button>
      </div>
    </div>
  );
};

export default ListView;
