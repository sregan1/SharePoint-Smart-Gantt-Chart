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
}

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

export type ViewMode = 'gantt' | 'list' | 'kanban' | 'dashboard';
export type ZoomLevel = 'day' | 'week' | 'month' | 'quarter';

// ─── Gantt display settings ───────────────────────────────────────────────────

export type GanttColorBy = 'status' | 'priority' | 'phase';
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
  showProgressText: boolean;
  showAssignee: boolean;
}

export const DEFAULT_GANTT_SETTINGS: IGanttDisplaySettings = {
  colorBy: 'phase',
  weekLabel: 'dates',
  headerTheme: 'dark',
  barStyle: 'gradient',
  rowHeight: 40,
  showWeekends: true,
  showDependencies: true,
  showProgressText: true,
  showAssignee: false,
};

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
