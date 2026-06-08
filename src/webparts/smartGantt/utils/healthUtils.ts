import { differenceInDays, parseISO, isValid } from 'date-fns';
import { ITask, IProject, TaskHealth, ProjectHealth } from '../models';

function parseDate(value: string): Date | null {
  if (!value) return null;
  const d = parseISO(value);
  return isValid(d) ? d : null;
}

export function computeTaskHealth(task: ITask, today: Date = new Date()): TaskHealth {
  if (task.status === 'Completed' || task.status === 'Cancelled') return 'complete';

  const due = parseDate(task.dueDate);
  const start = parseDate(task.startDate);

  if (due && due < today) return 'overdue';

  if (start && start < today && task.status === 'Not Started') return 'at-risk';

  if (start && due) {
    const totalDays = differenceInDays(due, start);
    if (totalDays >= 3) {
      const elapsedDays = differenceInDays(today, start);
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
