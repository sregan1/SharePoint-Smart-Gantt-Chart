import * as React from 'react';
import * as strings from 'SmartGanttWebPartStrings';
import { formatString } from '../localeUtils';
import {
  ITask, IProject, TaskStatus, TaskPriority, IWorkingCalendar, TaskHealth,
  STATUS_COLORS, STATUS_LIGHT_COLORS, PRIORITY_COLORS,
  TASK_STATUS_OPTIONS, TASK_PRIORITY_OPTIONS,
} from '../../models';
import { computeTaskHealth, hasDependencyViolation } from '../../utils/healthUtils';
import { formatDateOnly, parseDateOnly, toDateOnly } from '../../utils/dateUtils';
import { isWorkingDay } from '../../utils/scheduleUtils';
import { isOverdue, initials, stringToColor, getStatusLabel, getPriorityLabel } from '../../utils/taskDisplayUtils';
import { exportTasksCsv } from '../../services/ExportService';
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
  /** Deletes several tasks at once (the caller owns confirmation). The bulk
   *  "Delete selected" action is only offered when this is provided. */
  onBulkDelete?: (ids: number[]) => void;
  /** Working calendar — start/due dates on non-working days are marked. */
  calendar?: IWorkingCalendar;
}

// Worst health first when sorting ascending.
const HEALTH_SORT_ORDER: Record<TaskHealth, number> = { overdue: 0, 'at-risk': 1, 'on-track': 2, complete: 3 };

type SortField = 'sortOrder' | 'title' | 'startDate' | 'dueDate' | 'status' | 'priority' | 'assignedTo' | 'percentComplete' | 'phase' | 'health';
type SortDir = 'asc' | 'desc';

interface ISortThProps {
  field: SortField;
  label: string;
  width?: number;
  sortField: SortField;
  sortDir: SortDir;
  onSort: (field: SortField) => void;
}

// Module-scope (not defined inside ListView's render) so its component
// identity is stable across renders — defining it inline made React remount
// every header <th> (and its DOM/focus state) on every render.
const SortTh: React.FC<ISortThProps> = ({ field, label, width, sortField, sortDir, onSort }) => (
  <th
    className={sortField === field ? styles.sorted : ''}
    onClick={() => onSort(field)}
    onKeyDown={e => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onSort(field); }
    }}
    tabIndex={0}
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

// Keep % complete and status consistent when either is bulk-edited, matching
// the task panel's slider behavior.
function statusForPercent(task: ITask, v: number): TaskStatus | undefined {
  if (v === 100 && task.status !== 'Cancelled') return 'Completed';
  if (v > 0 && v < 100 && (task.status === 'Not Started' || task.status === 'Completed')) return 'In Progress';
  if (v === 0 && task.status === 'Completed') return 'Not Started';
  return undefined;
}

export const ListView: React.FC<IListViewProps> = ({
  tasks, project, showHealthBadges = true, onEditTask, onDeleteTask, onTaskUpdate, onAddTask, onBulkDelete, calendar,
}) => {
  const [sortField, setSortField] = React.useState<SortField>('sortOrder');
  const [sortDir, setSortDir] = React.useState<SortDir>('asc');
  const [selected, setSelected] = React.useState<Set<number>>(new Set());
  const [bulkAssignee, setBulkAssignee] = React.useState('');
  const [bulkPercent, setBulkPercent] = React.useState('');
  const selectAllRef = React.useRef<HTMLInputElement>(null);

  // Drop selections for tasks that are gone (deleted, or filtered out).
  React.useEffect(() => {
    setSelected(prev => {
      if (prev.size === 0) return prev;
      const ids = new Set(tasks.map(t => t.id));
      const next = new Set<number>();
      prev.forEach(id => { if (ids.has(id)) next.add(id); });
      return next.size === prev.size ? prev : next;
    });
  }, [tasks]);

  const allSelected = tasks.length > 0 && selected.size === tasks.length;
  React.useEffect(() => {
    if (selectAllRef.current) selectAllRef.current.indeterminate = selected.size > 0 && !allSelected;
  }, [selected, allSelected]);

  const toggleSelected = (id: number): void => {
    setSelected(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };

  const toggleAll = (): void => {
    setSelected(allSelected ? new Set() : new Set(tasks.map(t => t.id)));
  };

  // Bulk edits reuse the single-task update path, one call per selected task.
  const applyBulk = (updatesFor: (t: ITask) => Partial<ITask> | null): void => {
    tasks.forEach(t => {
      if (!selected.has(t.id)) return;
      const updates = updatesFor(t);
      if (updates) onTaskUpdate(t.id, updates);
    });
  };

  const bulkSetStatus = (status: TaskStatus): void => {
    applyBulk(t => {
      if (t.status === status) return null;
      const updates: Partial<ITask> = { status };
      if (status === 'Completed' && t.percentComplete < 100) updates.percentComplete = 100;
      if (status === 'Not Started' && t.percentComplete > 0) updates.percentComplete = 0;
      return updates;
    });
  };

  const bulkSetPriority = (priority: TaskPriority): void => {
    applyBulk(t => (t.priority === priority ? null : { priority }));
  };

  const bulkSetAssignee = (): void => {
    const name = bulkAssignee.trim();
    applyBulk(t => (t.assignedTo === name ? null : { assignedTo: name, assignedToEmail: '' }));
    setBulkAssignee('');
  };

  const bulkSetPercent = (): void => {
    const n = parseInt(bulkPercent, 10);
    if (isNaN(n)) return;
    const v = Math.min(100, Math.max(0, n));
    applyBulk(t => {
      const updates: Partial<ITask> = { percentComplete: v };
      const status = statusForPercent(t, v);
      if (status) updates.status = status;
      return updates;
    });
    setBulkPercent('');
  };

  const isNonWorking = (date: string): boolean => {
    if (!calendar || !date) return false;
    const d = toDateOnly(date);
    return !!d && !isWorkingDay(d, calendar);
  };

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
    tasks.forEach(t => { if (hasDependencyViolation(t, taskById)) set.add(t.id); });
    return set;
  }, [tasks, taskById]);

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

    // A missing date always sorts after every real date, regardless of
    // direction — otherwise ascending order put every undated task first.
    const UNDATED = Number.MAX_SAFE_INTEGER;

    const sortFn = (a: ITask, b: ITask): number => {
      let cmp: number;
      if (sortField === 'startDate' || sortField === 'dueDate') {
        const av = parseDateOnly(a[sortField])?.getTime() ?? UNDATED;
        const bv = parseDateOnly(b[sortField])?.getTime() ?? UNDATED;
        cmp = av - bv;
      } else if (sortField === 'health') {
        cmp = HEALTH_SORT_ORDER[computeTaskHealth(a)] - HEALTH_SORT_ORDER[computeTaskHealth(b)];
      } else if (sortField === 'percentComplete' || sortField === 'sortOrder') {
        cmp = a[sortField] - b[sortField];
      } else if (sortField === 'priority') {
        cmp = TASK_PRIORITY_OPTIONS.indexOf(a.priority) - TASK_PRIORITY_OPTIONS.indexOf(b.priority);
      } else if (sortField === 'status') {
        cmp = TASK_STATUS_OPTIONS.indexOf(a.status) - TASK_STATUS_OPTIONS.indexOf(b.status);
      } else {
        cmp = String(a[sortField] ?? '').localeCompare(String(b[sortField] ?? ''), undefined, { sensitivity: 'base' });
      }
      if (cmp === 0) cmp = a.id - b.id;
      // Undated tasks must land last in BOTH directions — negating cmp for
      // desc would otherwise put them first again.
      const isUndatedTiebreak = (sortField === 'startDate' || sortField === 'dueDate')
        && ((parseDateOnly(a[sortField]) === null) !== (parseDateOnly(b[sortField]) === null));
      if (isUndatedTiebreak) return parseDateOnly(a[sortField]) === null ? 1 : -1;
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

    // Recurse through every nesting level — a sub-task can itself have
    // sub-tasks (via a direct list edit or import), so only appending direct
    // children silently dropped grandchildren from the grid. `visited`
    // guards against a parent-cycle in the raw data looping forever.
    const pushTask = (t: ITask, visited: Set<number> = new Set()): void => {
      rows.push({ type: 'task', task: t, isChild: visited.size > 0 });
      if (visited.has(t.id)) return;
      visited.add(t.id);
      (children.get(t.id) || []).sort(sortFn).forEach(c => pushTask(c, visited));
    };

    byPhase.forEach((pTasks, phase) => {
      rows.push({ type: 'phase', phase });
      [...pTasks].sort(sortFn).forEach(t => pushTask(t));
    });
    [...noPhase].sort(sortFn).forEach(t => pushTask(t));

    return rows;
  }, [tasks, sortField, sortDir]);

  if (tasks.length === 0) {
    return (
      <div className={styles.listView}>
        <div className={styles.emptyState}>
          <div style={{ fontSize: 40, opacity: 0.3 }}>📋</div>
          <div style={{ fontSize: 16, fontWeight: 600, color: 'var(--neutralPrimary, #323130)' }}>{strings.ListView_EmptyTitle}</div>
          <button
            style={{
              background: 'var(--themePrimary, #0078D4)', color: '#fff', border: 'none', borderRadius: 4,
              padding: '8px 16px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
            }}
            onClick={onAddTask}
          >
            {strings.ListView_AddFirstTask}
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.listView}>
      {selected.size > 0 ? (
        <div className={styles.bulkBar} role="toolbar" aria-label={strings.View_List_BulkBarAriaLabel}>
          <span className={styles.bulkCount} role="status">{formatString(strings.View_List_BulkSelectedCount, { count: selected.size })}</span>
          <select
            className={styles.bulkControl}
            value=""
            aria-label={strings.View_List_BulkStatusPlaceholder}
            onChange={e => { if (e.target.value) bulkSetStatus(e.target.value as TaskStatus); }}
          >
            <option value="">{strings.View_List_BulkStatusPlaceholder}</option>
            {TASK_STATUS_OPTIONS.map(s => <option key={s} value={s}>{getStatusLabel(s)}</option>)}
          </select>
          <select
            className={styles.bulkControl}
            value=""
            aria-label={strings.View_List_BulkPriorityPlaceholder}
            onChange={e => { if (e.target.value) bulkSetPriority(e.target.value as TaskPriority); }}
          >
            <option value="">{strings.View_List_BulkPriorityPlaceholder}</option>
            {TASK_PRIORITY_OPTIONS.map(p => <option key={p} value={p}>{getPriorityLabel(p)}</option>)}
          </select>
          <input
            className={styles.bulkControl}
            style={{ width: 130 }}
            value={bulkAssignee}
            placeholder={strings.View_List_BulkAssigneePlaceholder}
            aria-label={strings.View_List_BulkAssigneePlaceholder}
            onChange={e => setBulkAssignee(e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter' && bulkAssignee.trim()) { e.preventDefault(); bulkSetAssignee(); } }}
          />
          <button className={styles.toolbarBtn} disabled={!bulkAssignee.trim()} onClick={bulkSetAssignee}>
            {strings.View_List_BulkAssignApply}
          </button>
          <input
            className={styles.bulkControl}
            style={{ width: 80 }}
            type="number"
            min={0}
            max={100}
            value={bulkPercent}
            placeholder={strings.View_List_BulkPercentPlaceholder}
            aria-label={strings.View_List_BulkPercentPlaceholder}
            onChange={e => setBulkPercent(e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); bulkSetPercent(); } }}
          />
          <button className={styles.toolbarBtn} disabled={bulkPercent === '' || isNaN(parseInt(bulkPercent, 10))} onClick={bulkSetPercent}>
            {strings.View_List_BulkPercentApply}
          </button>
          {onBulkDelete && (
            <button
              className={`${styles.toolbarBtn} ${styles.danger}`}
              onClick={() => { onBulkDelete(Array.from(selected)); setSelected(new Set()); }}
            >
              {strings.View_List_BulkDelete}
            </button>
          )}
          <button className={styles.toolbarBtn} onClick={() => setSelected(new Set())}>
            {strings.View_List_BulkClear}
          </button>
        </div>
      ) : (
        <div className={styles.listToolbar}>
          <button className={styles.toolbarBtn} onClick={() => exportTasksCsv(project, tasks)}>
            {strings.View_List_ExportCsv}
          </button>
        </div>
      )}
      <div className={styles.tableWrapper}>
        <table>
          <thead className={styles.thead}>
            <tr>
              <th className={styles.checkCol} style={{ width: 36, cursor: 'default' }}>
                <input
                  ref={selectAllRef}
                  type="checkbox"
                  checked={allSelected}
                  onChange={toggleAll}
                  aria-label={strings.View_List_SelectAllAriaLabel}
                />
              </th>
              <SortTh field="title" label={strings.ListView_ColTaskName} sortField={sortField} sortDir={sortDir} onSort={handleSort} />
              <SortTh field="status" label={strings.ListView_ColStatus} width={130} sortField={sortField} sortDir={sortDir} onSort={handleSort} />
              {showHealthBadges && <SortTh field="health" label={strings.ListView_ColHealth} width={100} sortField={sortField} sortDir={sortDir} onSort={handleSort} />}
              <SortTh field="priority" label={strings.ListView_ColPriority} width={100} sortField={sortField} sortDir={sortDir} onSort={handleSort} />
              <SortTh field="startDate" label={strings.ListView_ColStart} width={110} sortField={sortField} sortDir={sortDir} onSort={handleSort} />
              <SortTh field="dueDate" label={strings.ListView_ColDue} width={110} sortField={sortField} sortDir={sortDir} onSort={handleSort} />
              <SortTh field="assignedTo" label={strings.ListView_ColAssignedTo} width={140} sortField={sortField} sortDir={sortDir} onSort={handleSort} />
              <SortTh field="percentComplete" label={strings.ListView_ColProgress} width={140} sortField={sortField} sortDir={sortDir} onSort={handleSort} />
              <SortTh field="phase" label={strings.ListView_ColPhase} width={110} sortField={sortField} sortDir={sortDir} onSort={handleSort} />
              <th style={{ width: 160 }}>{strings.ListView_ColPredecessors}</th>
              <th style={{ width: 72 }} />
            </tr>
          </thead>
          <tbody className={styles.tbody}>
            {sortedTasks.map((row, _i) => {
              if (row.type === 'phase') {
                return (
                  <tr key={`phase-${row.phase}`} className={styles.phaseGroupRow}>
                    <td colSpan={showHealthBadges ? 12 : 11}>
                      <span className={styles.phaseGroupCell}>▸ {row.phase}</span>
                    </td>
                  </tr>
                );
              }

              const task = row.task!;
              const isChild = !!row.isChild;
              const overdue = isOverdue(task);

              return (
                <tr key={`task-${task.id}`} className={selected.has(task.id) ? styles.selectedRow : ''}>
                  {/* Select */}
                  <td className={styles.checkCol}>
                    <input
                      type="checkbox"
                      checked={selected.has(task.id)}
                      onChange={() => toggleSelected(task.id)}
                      aria-label={formatString(strings.View_List_SelectRowAriaLabel, { taskTitle: task.title })}
                    />
                  </td>

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
                          title={strings.ListView_DependencyViolationTooltip}
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
                      aria-label={formatString(strings.ListView_StatusAriaLabel, { taskTitle: task.title })}
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
                        <option key={s} value={s}>{getStatusLabel(s)}</option>
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
                      aria-label={formatString(strings.ListView_PriorityAriaLabel, { taskTitle: task.title })}
                      onChange={e => onTaskUpdate(task.id, { priority: e.target.value as TaskPriority })}
                      style={{
                        color: PRIORITY_COLORS[task.priority],
                        fontWeight: 600,
                      }}
                    >
                      {TASK_PRIORITY_OPTIONS.map(p => (
                        <option key={p} value={p}>{getPriorityLabel(p)}</option>
                      ))}
                    </select>
                  </td>

                  {/* Start */}
                  <td>
                    <span
                      className={`${styles.dateCell} ${isNonWorking(task.startDate) ? styles.nonWorking : ''}`}
                      title={isNonWorking(task.startDate) ? strings.View_List_NonWorkingDay : undefined}
                    >
                      {formatDateOnly(task.startDate, 'MMM d, yyyy')}
                    </span>
                  </td>

                  {/* Due */}
                  <td>
                    <span
                      className={`${styles.dateCell} ${overdue ? styles.overdue : ''} ${isNonWorking(task.dueDate) ? styles.nonWorking : ''}`}
                      title={isNonWorking(task.dueDate) ? strings.View_List_NonWorkingDay : undefined}
                    >
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
                      <span style={{ color: 'var(--neutralQuaternary, #C8C6C4)', fontSize: 12 }}>{strings.ListView_Unassigned}</span>
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
                    <span style={{ fontSize: 12, color: 'var(--neutralSecondary, #605E5C)' }}>
                      {task.phase || '—'}
                    </span>
                  </td>

                  {/* Predecessors */}
                  <td style={{ maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    <span style={{ fontSize: 12, color: 'var(--neutralSecondary, #605E5C)' }} title={task.dependencies.map(id => taskById.get(id)?.title ?? `#${id}`).join(', ')}>
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
                        title={strings.ListView_EditTitle}
                        aria-label={formatString(strings.ListView_EditTaskAriaLabel, { taskTitle: task.title })}
                      >
                        ✏
                      </button>
                      <button
                        className={`${styles.rowActionBtn} ${styles.deleteBtn}`}
                        onClick={() => onDeleteTask(task.id)}
                        title={strings.ListView_DeleteTitle}
                        aria-label={formatString(strings.ListView_DeleteTaskAriaLabel, { taskTitle: task.title })}
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
          {strings.ListView_AddTaskRow}
        </button>
      </div>
    </div>
  );
};

export default ListView;
