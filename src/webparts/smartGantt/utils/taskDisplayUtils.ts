import * as strings from 'SmartGanttWebPartStrings';
import { ITask, TaskStatus, TaskPriority, ProjectStatus } from '../models';
import { parseDateOnly, todayLocalMidnight } from './dateUtils';

/** True when a task's due date has passed and it isn't Completed/Cancelled. */
export function isOverdue(task: ITask): boolean {
  if (!task.dueDate || task.status === 'Completed' || task.status === 'Cancelled') return false;
  const due = parseDateOnly(task.dueDate);
  return !!due && due < todayLocalMidnight();
}

/** Up to two initials from a display name, for avatar circles. */
export function initials(name: string): string {
  if (!name) return '?';
  return name.split(' ').map(p => p[0]).join('').substring(0, 2).toUpperCase();
}

/** Deterministic HSL color from a name, so the same person's avatar always matches. */
export function stringToColor(s: string): string {
  let hash = 0;
  for (let i = 0; i < s.length; i++) hash = s.charCodeAt(i) + ((hash << 5) - hash);
  return `hsl(${Math.abs(hash) % 360}, 55%, 45%)`;
}

// Display-only label maps: the stored SharePoint Choice value (English) never
// changes, only what's rendered to the user.
const TASK_STATUS_LABELS: Record<TaskStatus, string> = {
  'Not Started': strings.Status_NotStarted,
  'In Progress': strings.Status_InProgress,
  'Completed': strings.Status_Completed,
  'On Hold': strings.Status_OnHold,
  'Cancelled': strings.Status_Cancelled,
};

const TASK_PRIORITY_LABELS: Record<TaskPriority, string> = {
  'Critical': strings.Priority_Critical,
  'High': strings.Priority_High,
  'Medium': strings.Priority_Medium,
  'Low': strings.Priority_Low,
};

const PROJECT_STATUS_LABELS: Record<ProjectStatus, string> = {
  'Planning': strings.ProjectStatus_Planning,
  'Active': strings.ProjectStatus_Active,
  'On Hold': strings.ProjectStatus_OnHold,
  'Completed': strings.ProjectStatus_Completed,
  'Cancelled': strings.ProjectStatus_Cancelled,
};

/** Localized display label for a task status (stored value is unchanged English). */
export function getStatusLabel(status: TaskStatus): string {
  return TASK_STATUS_LABELS[status] || status;
}

/** Localized display label for a task priority (stored value is unchanged English). */
export function getPriorityLabel(priority: TaskPriority): string {
  return TASK_PRIORITY_LABELS[priority] || priority;
}

/** Localized display label for a project status (stored value is unchanged English). */
export function getProjectStatusLabel(status: ProjectStatus): string {
  return PROJECT_STATUS_LABELS[status] || status;
}
