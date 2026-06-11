import { differenceInDays } from 'date-fns';
import { ITask, IProject, TaskHealth, ProjectHealth } from '../models';
import { parseDateOnly } from './dateUtils';

export function computeTaskHealth(task: ITask, today: Date = new Date()): TaskHealth {
  if (task.status === 'Completed' || task.status === 'Cancelled') return 'complete';

  // Compare calendar days, not instants: a task due today is not overdue.
  const todayDay = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const due = parseDateOnly(task.dueDate);
  const start = parseDateOnly(task.startDate);

  if (due && due < todayDay) return 'overdue';

  if (start && start < todayDay && task.status === 'Not Started') return 'at-risk';

  if (start && due) {
    const totalDays = differenceInDays(due, start);
    if (totalDays >= 3) {
      const elapsedDays = differenceInDays(todayDay, start);
      if (elapsedDays > 0) {
        const expectedPct = Math.min(100, (elapsedDays / totalDays) * 100);
        if (task.percentComplete < expectedPct - 10) return 'at-risk';
      }
    }
  }

  return 'on-track';
}

export function computeProjectHealth(tasks: ITask[], _project: IProject, today: Date = new Date()): ProjectHealth {
  if (tasks.length === 0) return 'on-track';

  const activeTasks = tasks.filter(t => t.status !== 'Completed' && t.status !== 'Cancelled');
  if (activeTasks.length === 0) return 'complete';

  const healths = tasks.map(t => computeTaskHealth(t, today));
  const taskHealthPairs = tasks.map((t, i) => ({ task: t, health: healths[i] }));

  const criticalOrHigh = taskHealthPairs.filter(
    p => p.task.priority === 'Critical' || p.task.priority === 'High'
  );

  if (criticalOrHigh.some(p => p.health === 'overdue')) return 'overdue';
  if (criticalOrHigh.some(p => p.health === 'at-risk')) return 'at-risk';
  if (taskHealthPairs.some(p => p.health === 'overdue')) return 'at-risk';
  if (taskHealthPairs.some(p => p.health === 'at-risk')) return 'at-risk';

  return 'on-track';
}

export function healthColor(h: TaskHealth | ProjectHealth): string {
  switch (h) {
    case 'complete':  return '#107C10';
    case 'on-track':  return '#0078D4';
    case 'at-risk':   return '#CA5010';
    case 'overdue':   return '#D13438';
  }
}

export function healthLightColor(h: TaskHealth | ProjectHealth): string {
  switch (h) {
    case 'complete':  return '#F1FAF1';
    case 'on-track':  return '#EFF6FC';
    case 'at-risk':   return '#FFF4EC';
    case 'overdue':   return '#FDF3F4';
  }
}

export function healthLabel(h: TaskHealth | ProjectHealth): string {
  switch (h) {
    case 'complete':  return 'Done';
    case 'on-track':  return 'On Track';
    case 'at-risk':   return 'At Risk';
    case 'overdue':   return 'Overdue';
  }
}

// ─── Critical path ────────────────────────────────────────────────────────────

/**
 * Computes the set of task ids on the critical path: the dependency chain with
 * the latest finish, walked backwards through each predecessor whose finish
 * immediately gates its successor. Tasks without dates are skipped.
 */
export function computeCriticalPath(tasks: ITask[]): Set<number> {
  const byId = new Map(tasks.map(t => [t.id, t]));
  const finish = new Map<number, number>(); // task id → critical finish time (ms)
  const visiting = new Set<number>();

  const getFinish = (t: ITask): number => {
    if (finish.has(t.id)) return finish.get(t.id)!;
    if (visiting.has(t.id)) return -Infinity; // dependency cycle — bail out
    visiting.add(t.id);
    const own = parseDateOnly(t.dueDate)?.getTime() ?? -Infinity;
    let result = own;
    for (const depId of t.dependencies) {
      const dep = byId.get(depId);
      if (dep) result = Math.max(result, getFinish(dep));
    }
    visiting.delete(t.id);
    finish.set(t.id, result);
    return result;
  };

  let endTask: ITask | null = null;
  let endFinish = -Infinity;
  for (const t of tasks) {
    if (t.status === 'Cancelled') continue;
    const f = getFinish(t);
    if (f > endFinish) { endFinish = f; endTask = t; }
  }

  const path = new Set<number>();
  let cur = endTask;
  while (cur && !path.has(cur.id)) {
    path.add(cur.id);
    let next: ITask | null = null;
    let nextFinish = -Infinity;
    for (const depId of cur.dependencies) {
      const dep = byId.get(depId);
      if (!dep) continue;
      const f = finish.get(dep.id) ?? -Infinity;
      if (f > nextFinish) { nextFinish = f; next = dep; }
    }
    cur = next;
  }
  return path;
}

/** True when a task is In Progress/Completed but a dependency is not yet complete. */
export function hasDependencyViolation(task: ITask, allTasks: ITask[]): boolean {
  if (task.status === 'Not Started' || task.status === 'Cancelled') return false;
  if (!task.dependencies.length) return false;
  const byId = new Map(allTasks.map(t => [t.id, t]));
  return task.dependencies.some(depId => {
    const dep = byId.get(depId);
    return !!dep && dep.status !== 'Completed' && dep.status !== 'Cancelled';
  });
}
