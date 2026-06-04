import * as React from 'react';
import { format } from 'date-fns';
import {
  ITask, IProject, TaskStatus, TaskPriority,
  STATUS_COLORS, STATUS_LIGHT_COLORS, PRIORITY_COLORS,
  TASK_STATUS_OPTIONS, TASK_PRIORITY_OPTIONS,
} from '../../models';
import styles from './ListView.module.scss';

interface IListViewProps {
  tasks: ITask[];
  project: IProject;
  onEditTask: (task: ITask) => void;
  onDeleteTask: (id: number) => void;
  onTaskUpdate: (id: number, updates: Partial<ITask>) => void;
  onAddTask: () => void;
}

type SortField = 'title' | 'startDate' | 'dueDate' | 'status' | 'priority' | 'assignedTo' | 'percentComplete' | 'phase';
type SortDir = 'asc' | 'desc';

function formatDate(s: string): string {
  if (!s) return '—';
  try {
    return format(new Date(s), 'MMM d, yyyy');
  } catch {
    return '—';
  }
}

function isOverdue(task: ITask): boolean {
  if (!task.dueDate || task.status === 'Completed' || task.status === 'Cancelled') return false;
  return new Date(task.dueDate) < new Date();
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
  tasks, onEditTask, onDeleteTask, onTaskUpdate, onAddTask,
}) => {
  const [sortField, setSortField] = React.useState<SortField>('sortOrder' as any);
  const [sortDir, setSortDir] = React.useState<SortDir>('asc');

  const handleSort = (field: SortField): void => {
    if (sortField === field) {
      setSortDir(d => d === 'asc' ? 'desc' : 'asc');
    } else {
      setSortField(field);
      setSortDir('asc');
    }
  };

  const sortedTasks = React.useMemo(() => {
    const top = tasks.filter(t => !t.parentTaskId);
    const children = new Map<number, ITask[]>();
    tasks.filter(t => t.parentTaskId).forEach(t => {
      if (!children.has(t.parentTaskId!)) children.set(t.parentTaskId!, []);
      children.get(t.parentTaskId!)!.push(t);
    });

    const sortFn = (a: ITask, b: ITask): number => {
      let av: any = (a as any)[sortField] || '';
      let bv: any = (b as any)[sortField] || '';
      if (sortField === 'startDate' || sortField === 'dueDate') {
        av = av ? new Date(av).getTime() : 0;
        bv = bv ? new Date(bv).getTime() : 0;
      }
      if (av < bv) return sortDir === 'asc' ? -1 : 1;
      if (av > bv) return sortDir === 'asc' ? 1 : -1;
      return 0;
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

    const rows: Array<{ type: 'task' | 'phase'; task?: ITask; phase?: string }> = [];

    byPhase.forEach((pTasks, phase) => {
      rows.push({ type: 'phase', phase });
      [...pTasks].sort(sortFn).forEach(t => {
        rows.push({ type: 'task', task: t });
        (children.get(t.id) || []).sort(sortFn).forEach(c => rows.push({ type: 'task', task: c }));
      });
    });
    [...noPhase].sort(sortFn).forEach(t => {
      rows.push({ type: 'task', task: t });
      (children.get(t.id) || []).sort(sortFn).forEach(c => rows.push({ type: 'task', task: c }));
    });

    return rows;
  }, [tasks, sortField, sortDir]);

  const SortTh: React.FC<{ field: SortField; label: string; width?: number }> = ({ field, label, width }) => (
    <th
      className={sortField === field ? styles.sorted : ''}
      onClick={() => handleSort(field)}
      style={width ? { width } : undefined}
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
              <SortTh field="priority" label="Priority" width={100} />
              <SortTh field="startDate" label="Start" width={110} />
              <SortTh field="dueDate" label="Due" width={110} />
              <SortTh field="assignedTo" label="Assigned To" width={140} />
              <SortTh field="percentComplete" label="Progress" width={140} />
              <SortTh field="phase" label="Phase" width={110} />
              <th style={{ width: 72 }} />
            </tr>
          </thead>
          <tbody className={styles.tbody}>
            {sortedTasks.map((row, _i) => {
              if (row.type === 'phase') {
                return (
                  <tr key={`phase-${row.phase}`} className={styles.phaseGroupRow}>
                    <td colSpan={9}>
                      <span className={styles.phaseGroupCell}>▸ {row.phase}</span>
                    </td>
                  </tr>
                );
              }

              const task = row.task!;
              const isChild = !!task.parentTaskId;
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
                    </div>
                  </td>

                  {/* Status */}
                  <td>
                    <select
                      className={styles.inlineSelect}
                      value={task.status}
                      onChange={e => onTaskUpdate(task.id, { status: e.target.value as TaskStatus })}
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

                  {/* Priority */}
                  <td>
                    <select
                      className={styles.inlineSelect}
                      value={task.priority}
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
                    <span className={styles.dateCell}>{formatDate(task.startDate)}</span>
                  </td>

                  {/* Due */}
                  <td>
                    <span className={`${styles.dateCell} ${overdue ? styles.overdue : ''}`}>
                      {formatDate(task.dueDate)}
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

                  {/* Actions */}
                  <td>
                    <div className={styles.rowActions}>
                      <button
                        className={styles.rowActionBtn}
                        onClick={() => onEditTask(task)}
                        title="Edit"
                      >
                        ✏
                      </button>
                      <button
                        className={`${styles.rowActionBtn} ${styles.deleteBtn}`}
                        onClick={() => onDeleteTask(task.id)}
                        title="Delete"
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
