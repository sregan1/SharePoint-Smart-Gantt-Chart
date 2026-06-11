import * as XLSX from 'xlsx';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ITask, TaskStatus, TaskPriority, TASK_STATUS_OPTIONS, TASK_PRIORITY_OPTIONS } from '../models';
import { SharePointService } from './SharePointService';
import { toDateOnly } from '../utils/dateUtils';

// ─── Public types ─────────────────────────────────────────────────────────────

export type ImportableField = keyof Pick<
  ITask,
  | 'title' | 'startDate' | 'dueDate' | 'status' | 'priority'
  | 'assignedTo' | 'assignedToEmail' | 'percentComplete'
  | 'phase' | 'description' | 'notes' | 'isMilestone'
> | 'skip';

export interface IImportableFieldDef {
  key: ImportableField;
  label: string;
  required?: boolean;
}

export const IMPORTABLE_FIELDS: IImportableFieldDef[] = [
  { key: 'title', label: 'Task Name', required: true },
  { key: 'startDate', label: 'Start Date' },
  { key: 'dueDate', label: 'Due Date' },
  { key: 'status', label: 'Status' },
  { key: 'priority', label: 'Priority' },
  { key: 'assignedTo', label: 'Assigned To' },
  { key: 'assignedToEmail', label: 'Assigned To (Email)' },
  { key: 'percentComplete', label: '% Complete' },
  { key: 'phase', label: 'Phase' },
  { key: 'description', label: 'Description' },
  { key: 'notes', label: 'Notes' },
  { key: 'isMilestone', label: 'Is Milestone' },
  { key: 'skip', label: 'Skip this column' },
];

export type ColumnMapping = Record<string, ImportableField>;

export interface IImportSource {
  type: 'excel' | 'planner';
  fileName?: string;           // Excel only
  planId?: string;             // Planner only
  planName?: string;           // Planner only
  headers: string[];
  rows: Record<string, string>[];
  autoMapping: ColumnMapping;
  needsMapping: boolean;
}

export interface IPlannerPlan {
  id: string;
  title: string;
  groupId: string;
  groupName: string;
}

export interface IBatchImportResult {
  succeeded: number;
  failed: number;
  errors: string[];
}

// ─── Auto-map aliases ─────────────────────────────────────────────────────────

const ALIASES: Record<string, ImportableField> = {};
const registerAlias = (aliases: string[], field: ImportableField): void => {
  aliases.forEach(a => { ALIASES[a.toLowerCase()] = field; });
};

registerAlias(['title', 'task', 'task name', 'name', 'subject', 'summary', 'work item'], 'title');
registerAlias(['start', 'start date', 'begin', 'begin date', 'started', 'start_date', 'startdate'], 'startDate');
registerAlias(['end', 'finish', 'due', 'due date', 'deadline', 'end date', 'finish date', 'due_date', 'duedate'], 'dueDate');
registerAlias(['status', 'state', 'task status', 'progress status'], 'status');
registerAlias(['priority', 'urgency', 'importance', 'severity'], 'priority');
registerAlias(['assigned to', 'owner', 'responsible', 'resource', 'assignee', 'assigned', 'assigned_to', 'reporter'], 'assignedTo');
registerAlias(['email', 'assigned email', 'owner email', 'user email', 'resource email'], 'assignedToEmail');
registerAlias(['% complete', 'percent complete', '% done', 'percent', 'completion', 'progress', 'done %', 'complete', 'completion %'], 'percentComplete');
registerAlias(['phase', 'category', 'group', 'bucket', 'sprint', 'iteration', 'epic', 'module', 'section'], 'phase');
registerAlias(['description', 'task description', 'detail', 'details', 'desc'], 'description');
registerAlias(['notes', 'comments', 'comment', 'remarks', 'note', 'annotation'], 'notes');
registerAlias(['milestone', 'is milestone', 'key milestone', 'milestone?', 'ismilestone'], 'isMilestone');

function autoMap(headers: string[]): ColumnMapping {
  const mapping: ColumnMapping = {};
  const usedFields = new Set<string>();

  headers.forEach(header => {
    const key = header.toLowerCase().trim();
    const field = ALIASES[key];
    if (field && field !== 'skip' && !usedFields.has(field)) {
      mapping[header] = field;
      usedFields.add(field);
    } else {
      mapping[header] = 'skip';
    }
  });
  return mapping;
}

function mappingNeedsReview(mapping: ColumnMapping): boolean {
  const mapped = Object.values(mapping).filter(v => v !== 'skip');
  const hasTitle = mapped.includes('title');
  const hasSkips = Object.values(mapping).some(v => v === 'skip');
  return !hasTitle || hasSkips;
}

// ─── Date normalization ───────────────────────────────────────────────────────

// All schedule dates are normalized to date-only 'YYYY-MM-DD' strings, the
// canonical form the rest of the app uses (see utils/dateUtils.ts).
function ymd(y: number, m1: number, d: number): string {
  return `${y}-${String(m1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

function parseExcelDate(value: string | number | null | undefined): string {
  if (!value && value !== 0) return '';

  if (typeof value === 'number') {
    // Excel serial date — already a pure calendar day
    const date = XLSX.SSF.parse_date_code(value);
    return date ? ymd(date.y, date.m, date.d) : '';
  }

  const str = String(value).trim();
  if (!str) return '';

  // ISO format (YYYY-MM-DD with optional time) — defer to the shared
  // normalizer, which handles UTC-midnight and legacy local-midnight values.
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    return toDateOnly(str);
  }

  // MM/DD/YYYY or M/D/YYYY — interpret as a calendar day directly
  const mdy = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (mdy) return ymd(+mdy[3], +mdy[1], +mdy[2]);

  // DD-MM-YYYY (dash-separated, day-first)
  const dmy = str.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (dmy) return ymd(+dmy[3], +dmy[2], +dmy[1]);

  // Fallback
  const d = new Date(str);
  return isNaN(d.getTime()) ? '' : toDateOnly(d.toISOString());
}

function normalizeStatus(raw: string): TaskStatus {
  const s = raw.toLowerCase().trim();
  if (['done', 'complete', 'completed', 'finished', 'closed', '100%'].includes(s)) return 'Completed';
  if (['in progress', 'in-progress', 'active', 'started', 'wip', 'doing'].includes(s)) return 'In Progress';
  if (['on hold', 'blocked', 'paused', 'deferred', 'waiting'].includes(s)) return 'On Hold';
  if (['cancelled', 'canceled', 'rejected', 'removed'].includes(s)) return 'Cancelled';
  return 'Not Started';
}

function normalizeStatus2(raw: string): TaskStatus | undefined {
  const cleaned = raw.toLowerCase().trim();
  const match = TASK_STATUS_OPTIONS.find(s => s.toLowerCase() === cleaned);
  if (match) return match;
  return normalizeStatus(cleaned);
}

function normalizePriority(raw: string | number): TaskPriority {
  if (typeof raw === 'number') {
    if (raw <= 1) return 'Critical';
    if (raw <= 4) return 'High';
    if (raw <= 7) return 'Medium';
    return 'Low';
  }
  const s = String(raw).toLowerCase().trim();
  if (['critical', 'urgent', '1', 'p1', 'p0'].includes(s)) return 'Critical';
  if (['high', 'important', '2', 'p2'].includes(s)) return 'High';
  if (['medium', 'normal', 'moderate', '3', 'p3', 'mid'].includes(s)) return 'Medium';
  if (['low', 'minor', '4', 'p4'].includes(s)) return 'Low';
  const match = TASK_PRIORITY_OPTIONS.find(p => p.toLowerCase() === s);
  return match || 'Medium';
}

function normalizeBoolean(raw: string): boolean {
  const s = String(raw).toLowerCase().trim();
  return ['true', 'yes', '1', 'x', '✓', 'milestone'].includes(s);
}

// ─── Excel parsing ────────────────────────────────────────────────────────────

export async function parseExcelFile(file: File): Promise<IImportSource> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target!.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellDates: false });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawRows: Record<string, unknown>[] = XLSX.utils.sheet_to_json(sheet, {
          defval: '',
          raw: true,
        });

        if (rawRows.length === 0) {
          reject(new Error('The file appears to be empty or has no data rows.'));
          return;
        }

        const headers = Object.keys(rawRows[0]);
        const rows = rawRows.map(r => {
          const cleaned: Record<string, string> = {};
          headers.forEach(h => { cleaned[h] = String(r[h] ?? ''); });
          return cleaned;
        });

        const autoMapping = autoMap(headers);
        resolve({
          type: 'excel',
          fileName: file.name,
          headers,
          rows,
          autoMapping,
          needsMapping: mappingNeedsReview(autoMapping),
        });
      } catch {
        reject(new Error('Could not read the file. Make sure it is a valid .xlsx, .xls, or .csv file.'));
      }
    };
    reader.onerror = () => reject(new Error('File read error.'));
    reader.readAsArrayBuffer(file);
  });
}

// ─── Apply mapping & normalize ────────────────────────────────────────────────

export function applyMapping(
  rows: Record<string, string>[],
  mapping: ColumnMapping
): Partial<ITask>[] {
  return rows
    .map(row => {
      const task: Partial<ITask> = {};
      Object.entries(mapping).forEach(([col, field]) => {
        if (field === 'skip') return;
        const raw = row[col] ?? '';
        switch (field) {
          case 'title': task.title = raw; break;
          case 'startDate': task.startDate = parseExcelDate(raw); break;
          case 'dueDate': task.dueDate = parseExcelDate(raw); break;
          case 'status': task.status = raw ? normalizeStatus2(raw) : undefined; break;
          case 'priority': task.priority = raw ? normalizePriority(raw) : undefined; break;
          case 'assignedTo': task.assignedTo = raw; break;
          case 'assignedToEmail': task.assignedToEmail = raw; break;
          case 'percentComplete': {
            const n = parseFloat(raw.replace('%', ''));
            if (!isNaN(n)) task.percentComplete = Math.min(100, Math.max(0, n));
            break;
          }
          case 'phase': task.phase = raw; break;
          case 'description': task.description = raw; break;
          case 'notes': task.notes = raw; break;
          case 'isMilestone': task.isMilestone = raw ? normalizeBoolean(raw) : false; break;
        }
      });
      return task;
    })
    .filter(t => !!(t.title?.trim()));
}

// ─── Planner Graph integration ────────────────────────────────────────────────

interface IGraphClient {
  api(url: string): IGraphRequest;
}
interface IGraphRequest {
  filter(f: string): this;
  select(s: string): this;
  top(n: number): this;
  get(): Promise<{ value: any[] } | any>;
}

async function getGraphClient(context: WebPartContext): Promise<IGraphClient> {
  return (context as any).msGraphClientFactory.getClient('3') as IGraphClient;
}

export async function fetchPlannerPlans(context: WebPartContext): Promise<IPlannerPlan[]> {
  const graph = await getGraphClient(context);

  // Get the user's M365 groups
  let groups: any[] = [];
  try {
    const resp = await graph
      .api('/me/memberOf/microsoft.graph.group')
      .select('id,displayName,groupTypes')
      .top(50)
      .get();
    groups = resp.value || [];
  } catch {
    return [];
  }

  // Filter to M365 unified groups (Teams / Planner-capable)
  const m365Groups = groups.filter(
    (g: any) => Array.isArray(g.groupTypes) && g.groupTypes.includes('Unified')
  );

  const plans: IPlannerPlan[] = [];
  await Promise.all(
    m365Groups.map(async (group: any) => {
      try {
        const resp = await graph.api(`/groups/${group.id}/planner/plans`).get();
        const groupPlans: any[] = resp.value || [];
        groupPlans.forEach(p => {
          plans.push({
            id: p.id,
            title: p.title,
            groupId: group.id,
            groupName: group.displayName,
          });
        });
      } catch {
        // Group may not have Planner — skip silently
      }
    })
  );

  return plans.sort((a, b) => a.title.localeCompare(b.title));
}

export async function fetchPlannerTasks(
  context: WebPartContext,
  planId: string,
  planName: string
): Promise<IImportSource> {
  const graph = await getGraphClient(context);

  // Fetch tasks and buckets in parallel
  const [tasksResp, bucketsResp] = await Promise.all([
    graph.api(`/planner/plans/${planId}/tasks`).top(500).get(),
    graph.api(`/planner/plans/${planId}/buckets`).get(),
  ]);

  const plannerTasks: any[] = tasksResp.value || [];
  const buckets: any[] = bucketsResp.value || [];
  const bucketMap = new Map<string, string>(buckets.map((b: any) => [b.id, b.name]));

  // Collect unique user IDs from assignments to resolve names
  const userIds = new Set<string>();
  plannerTasks.forEach((t: any) => {
    if (t.assignments) Object.keys(t.assignments).forEach(uid => userIds.add(uid));
  });

  const userMap = new Map<string, { name: string; email: string }>();
  await Promise.all(
    Array.from(userIds).map(async uid => {
      try {
        const user = await graph.api(`/users/${uid}`).select('displayName,mail').get();
        userMap.set(uid, { name: user.displayName || '', email: user.mail || '' });
      } catch {
        userMap.set(uid, { name: uid, email: '' });
      }
    })
  );

  // Planner priority scale: 1 = Urgent, 3 = Important, 5 = Medium, 9 = Low.
  // Map to labels here — the generic string normalizer would misread "3".
  const plannerPriorityLabel = (p: number): string => {
    if (p <= 2) return 'Critical';
    if (p <= 4) return 'High';
    if (p <= 7) return 'Medium';
    return 'Low';
  };

  // Map Planner tasks to our row format
  const rows = plannerTasks.map((t: any) => {
    const assignedUserIds = t.assignments ? Object.keys(t.assignments) : [];
    const firstUser = assignedUserIds.length > 0 ? userMap.get(assignedUserIds[0]) : null;

    return {
      'Title': t.title || '',
      'Start Date': t.startDateTime ? toDateOnly(t.startDateTime) : '',
      'Due Date': t.dueDateTime ? toDateOnly(t.dueDateTime) : '',
      'Status': t.percentComplete === 100 ? 'Completed' : t.percentComplete > 0 ? 'In Progress' : 'Not Started',
      'Priority': plannerPriorityLabel(typeof t.priority === 'number' ? t.priority : 5),
      'Assigned To': firstUser?.name || '',
      'Assigned To (Email)': firstUser?.email || '',
      '% Complete': String(t.percentComplete ?? 0),
      'Phase': bucketMap.get(t.bucketId) || '',
    };
  });

  const headers = ['Title', 'Start Date', 'Due Date', 'Status', 'Priority', 'Assigned To', 'Assigned To (Email)', '% Complete', 'Phase'];
  const autoMapping = autoMap(headers);

  return {
    type: 'planner',
    planId,
    planName,
    headers,
    rows,
    autoMapping,
    needsMapping: false, // Planner fields are always well-known
  };
}

// ─── Batch import ─────────────────────────────────────────────────────────────

export async function batchImport(
  spService: SharePointService,
  listName: string,
  tasks: Partial<ITask>[],
  onProgress?: (done: number, total: number) => void
): Promise<IBatchImportResult> {
  const result: IBatchImportResult = { succeeded: 0, failed: 0, errors: [] };
  const total = tasks.length;

  for (let i = 0; i < tasks.length; i++) {
    const task = {
      ...tasks[i],
      status: tasks[i].status || 'Not Started',
      priority: tasks[i].priority || 'Medium',
      percentComplete: tasks[i].percentComplete ?? 0,
      sortOrder: i,
    };
    try {
      await spService.createTask(listName, task);
      result.succeeded++;
    } catch (err: any) {
      result.failed++;
      result.errors.push(`Row ${i + 1} ("${task.title}"): ${err?.message || 'Unknown error'}`);
    }
    if (onProgress) onProgress(i + 1, total);
  }

  return result;
}
