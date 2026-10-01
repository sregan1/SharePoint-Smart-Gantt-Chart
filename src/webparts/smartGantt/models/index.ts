export interface IProject {
  id: number;
  title: string;
  listName: string;
  description: string;
  color: string;
  startDate: string;
  dueDate: string;
  status: ProjectStatus;
  projectManager: string;
  projectManagerEmail: string;
  created: string;
  isArchived?: boolean;
  /** SharePoint concurrency token from the last read — passed back on update
   *  so a stale edit is rejected instead of silently overwriting a concurrent
   *  change. Absent for a project that hasn't been read from SharePoint yet. */
  etag?: string;
}

export type ProjectStatus = 'Planning' | 'Active' | 'On Hold' | 'Completed' | 'Cancelled';

export interface ITask {
  id: number;
  title: string;
  description: string;
  startDate: string;
  dueDate: string;
  status: TaskStatus;
  priority: TaskPriority;
  assignedTo: string;
  assignedToEmail: string;
  percentComplete: number;
  parentTaskId: number | null;
  dependencies: number[];
  notes: string;
  color: string;
  sortOrder: number;
  isMilestone: boolean;
  phase: string;
  created: string;
  modified: string;
  /** SharePoint concurrency token from the last read — see IProject.etag. */
  etag?: string;
  /** Per-predecessor link type and lag. Keys are predecessor ids that also
   *  appear in `dependencies`; a missing key means finish-to-start, lag 0.
   *  Stored in the Dependencies column as "12,15SS+2,18FF-1". */
  dependencyLinks?: Record<number, IDependencyLink>;
  /** Planned (baseline) dates, ISO yyyy-MM-dd, captured on demand. */
  baselineStart?: string;
  baselineDue?: string;
}

export type DependencyType = 'FS' | 'SS' | 'FF' | 'SF';

export interface IDependencyLink {
  type: DependencyType;
  /** Lag (positive) or lead (negative) in working days. */
  lag: number;
}

/** Project working calendar, supplied by web part properties. */
export interface IWorkingCalendar {
  /** Working weekdays, 0 = Sunday … 6 = Saturday. */
  workingDays: number[];
  /** Non-working dates, ISO yyyy-MM-dd. */
  holidays: string[];
}

export const DEFAULT_WORKING_CALENDAR: IWorkingCalendar = {
  workingDays: [1, 2, 3, 4, 5],
  holidays: [],
};

export type TaskStatus = 'Not Started' | 'In Progress' | 'Completed' | 'On Hold' | 'Cancelled';
export type TaskPriority = 'Critical' | 'High' | 'Medium' | 'Low';

export const STATUS_COLORS: Record<TaskStatus, string> = {
  'Not Started': '#8B929A',
  'In Progress': '#0078D4',
  'Completed': '#107C10',
  'On Hold': '#CA5010',
  'Cancelled': '#D13438',
};

export const STATUS_LIGHT_COLORS: Record<TaskStatus, string> = {
  'Not Started': '#F3F2F1',
  'In Progress': '#EFF6FC',
  'Completed': '#F1FAF1',
  'On Hold': '#FFF4EC',
  'Cancelled': '#FDF3F4',
};

// Project status uses its own palette — ProjectStatus and TaskStatus share
// some labels (Completed, On Hold, Cancelled) but not all (Planning, Active
// have no TaskStatus equivalent), so they can't share STATUS_COLORS without
// those two falling back to gray.
export const PROJECT_STATUS_COLORS: Record<ProjectStatus, string> = {
  Planning: '#8764B8',
  Active: '#0078D4',
  'On Hold': '#CA5010',
  Completed: '#107C10',
  Cancelled: '#8B929A',
};

export const PROJECT_STATUS_LIGHT_COLORS: Record<ProjectStatus, string> = {
  Planning: '#F3EFF8',
  Active: '#EFF6FC',
  'On Hold': '#FFF4EC',
  Completed: '#F1FAF1',
  Cancelled: '#F3F2F1',
};

export const PRIORITY_COLORS: Record<TaskPriority, string> = {
  Critical: '#D13438',
  High: '#CA5010',
  Medium: '#0078D4',
  Low: '#107C10',
};

export const PROJECT_STATUS_OPTIONS: ProjectStatus[] = [
  'Planning', 'Active', 'On Hold', 'Completed', 'Cancelled',
];

export const TASK_STATUS_OPTIONS: TaskStatus[] = [
  'Not Started', 'In Progress', 'Completed', 'On Hold', 'Cancelled',
];

export const TASK_PRIORITY_OPTIONS: TaskPriority[] = [
  'Critical', 'High', 'Medium', 'Low',
];

export const PROJECT_COLORS = [
  '#0078D4', '#107C10', '#CA5010', '#8764B8', '#038387',
  '#D13438', '#C43148', '#00B7C3', '#881798', '#498205',
];

export type ViewMode = 'gantt' | 'list' | 'kanban' | 'dashboard' | 'portfolio';
export type ZoomLevel = 'day' | 'week' | 'month' | 'quarter';

export type TaskHealth = 'complete' | 'on-track' | 'at-risk' | 'overdue';
export type ProjectHealth = 'complete' | 'on-track' | 'at-risk' | 'overdue';

export interface IProjectTaskStats {
  listName: string;
  totalTasks: number;
  byStatus: Record<TaskStatus, number>;
  overallPct: number;
  health: ProjectHealth;
  overdueCount: number;
  atRiskCount: number;
  inProgressCount: number;
  completedCount: number;
  milestoneCount: number;
  earliestStart: string;
  latestDue: string;
  /** Stats could not be loaded (permissions/throttling/network) — render as
   *  "unavailable" rather than a healthy empty project. */
  statsError?: boolean;
}

// ─── Gantt display settings ───────────────────────────────────────────────────

export type GanttColorBy = 'status' | 'priority' | 'phase' | 'health';
export type GanttWeekLabel = 'dates' | 'project' | 'iso';
export type GanttHeaderTheme = 'dark' | 'navy' | 'teal' | 'purple' | 'light';
export type GanttBarStyle = 'gradient' | 'flat';

export interface IGanttDisplaySettings {
  colorBy: GanttColorBy;
  weekLabel: GanttWeekLabel;
  headerTheme: GanttHeaderTheme;
  barStyle: GanttBarStyle;
  rowHeight: number;
  showWeekends: boolean;
  showDependencies: boolean;
  showCriticalPathOnly: boolean;
  dependenciesOnHover: boolean;
  showProgressText: boolean;
  showAssignee: boolean;
  showHealthBadges: boolean;
  showCriticalPath: boolean;
}

export const DEFAULT_GANTT_SETTINGS: IGanttDisplaySettings = {
  colorBy: 'phase',
  weekLabel: 'dates',
  headerTheme: 'dark',
  barStyle: 'gradient',
  rowHeight: 40,
  showWeekends: true,
  showDependencies: true,
  showCriticalPathOnly: true,
  dependenciesOnHover: true,
  showProgressText: true,
  showAssignee: false,
  showHealthBadges: true,
  showCriticalPath: false,
};

// ─── Task filtering (shared across Gantt / List / Kanban / Dashboard) ─────────

export type DueFilter = 'all' | 'overdue' | 'today' | 'week';

export interface ITaskFilter {
  text: string;
  statuses: TaskStatus[];
  priorities: TaskPriority[];
  assignees: string[];
  phases: string[];
  due: DueFilter;
}

export const EMPTY_TASK_FILTER: ITaskFilter = {
  text: '', statuses: [], priorities: [], assignees: [], phases: [], due: 'all',
};

export function isFilterActive(f: ITaskFilter): boolean {
  return !!f.text || f.statuses.length > 0 || f.priorities.length > 0
    || f.assignees.length > 0 || f.phases.length > 0 || f.due !== 'all';
}

export const HEADER_THEME_COLORS: Record<GanttHeaderTheme, { bg: string; text: string; subtext: string }> = {
  dark:   { bg: '#1B1B3A', text: 'rgba(255,255,255,0.9)', subtext: 'rgba(255,255,255,0.55)' },
  navy:   { bg: '#1565C0', text: 'rgba(255,255,255,0.95)', subtext: 'rgba(255,255,255,0.6)' },
  teal:   { bg: '#00695C', text: 'rgba(255,255,255,0.95)', subtext: 'rgba(255,255,255,0.6)' },
  purple: { bg: '#4527A0', text: 'rgba(255,255,255,0.95)', subtext: 'rgba(255,255,255,0.6)' },
  light:  { bg: '#F3F2F1', text: '#323130', subtext: '#605E5C' },
};

// Color palette for phase-based coloring (consistent hash → color)
export const PHASE_PALETTE = [
  '#0078D4', '#107C10', '#CA5010', '#8764B8', '#038387',
  '#D13438', '#00B7C3', '#881798', '#498205', '#C43148',
  '#005A9E', '#217346', '#8E562E', '#6B69D6', '#00AD56',
];

export function phaseColor(phase: string): string {
  let hash = 0;
  for (let i = 0; i < phase.length; i++) hash = phase.charCodeAt(i) + ((hash << 5) - hash);
  return PHASE_PALETTE[Math.abs(hash) % PHASE_PALETTE.length];
}
