import * as XLSX from 'xlsx';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ITask, TaskStatus, TaskPriority, TASK_STATUS_OPTIONS, TASK_PRIORITY_OPTIONS } from '../models';
import { SharePointService } from './SharePointService';
import { toDateOnly } from '../utils/dateUtils';
import * as strings from 'SmartGanttWebPartStrings';
import { formatString } from '../components/localeUtils';

// ─── Public types ─────────────────────────────────────────────────────────────

export type ImportableField = keyof Pick<
  ITask,
  | 'title' | 'startDate' | 'dueDate' | 'status' | 'priority'
  | 'assignedTo' | 'assignedToEmail' | 'percentComplete'
  | 'phase' | 'description' | 'notes' | 'isMilestone' | 'dependencies'
> | 'skip';

export interface IImportableFieldDef {
  key: ImportableField;
  label: string;
  required?: boolean;
}

export const IMPORTABLE_FIELDS: IImportableFieldDef[] = [
  { key: 'title', label: strings.ImportField_TaskName, required: true },
  { key: 'startDate', label: strings.ImportField_StartDate },
  { key: 'dueDate', label: strings.ImportField_DueDate },
  { key: 'status', label: strings.ImportField_Status },
  { key: 'priority', label: strings.ImportField_Priority },
  { key: 'assignedTo', label: strings.ImportField_AssignedTo },
  { key: 'assignedToEmail', label: strings.ImportField_AssignedToEmail },
  { key: 'percentComplete', label: strings.ImportField_PercentComplete },
  { key: 'phase', label: strings.ImportField_Phase },
  { key: 'description', label: strings.ImportField_Description },
  { key: 'notes', label: strings.ImportField_Notes },
  { key: 'isMilestone', label: strings.ImportField_IsMilestone },
  { key: 'dependencies', label: strings.ImportField_Dependencies },
  { key: 'skip', label: strings.ImportField_SkipColumn },
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
  /** SharePoint item id created for each input task, in the same order (null where creation failed). */
  createdIds: Array<number | null>;
}

export interface IResolveDependenciesResult {
  resolved: number;
  warnings: string[];
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
registerAlias(['dependencies', 'depends on', 'predecessors', 'predecessor', 'prerequisite', 'prerequisites', 'blockers', 'blocked by'], 'dependencies');

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

  // Excel serial date arriving as a stringified number (sheet_to_json with
  // raw: true stringifies every cell before this function sees it). Guard
  // the range so a plain numeric ID/row-number column isn't misread as a date.
  if (/^\d+(\.\d+)?$/.test(str)) {
    const serial = parseFloat(str);
    if (serial > 59 && serial < 200000) {
      const date = XLSX.SSF.parse_date_code(serial);
      if (date) return ymd(date.y, date.m, date.d);
    }
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
          reject(new Error(strings.ImportService_EmptyFileError));
          return;
        }

        const headers = Object.keys(rawRows[0]);
        const rows = rawRows.map(r => {
          const cleaned: Record<string, string> = {};
          headers.forEach(h => { cleaned[h] = String(r[h] ?? ''); });
          return cleaned;
        });

        const autoMapping = autoMap(headers);
        // Blank/merged header cells come back from SheetJS as '__EMPTY',
        // '__EMPTY_1', etc. — auto-skip them so they don't clutter the
        // mapping UI as columns needing review.
        headers.forEach(h => {
          if (/^__EMPTY(_\d+)?$/.test(h)) autoMapping[h] = 'skip';
        });
        resolve({
          type: 'excel',
          fileName: file.name,
          headers,
          rows,
          autoMapping,
          needsMapping: mappingNeedsReview(autoMapping),
        });
      } catch (e) {
        const detail = e instanceof Error ? e.message : '';
        reject(new Error(formatString(strings.ImportService_ReadErrorDetail, { detail })));
      }
    };
    reader.onerror = () => reject(new Error(strings.ImportService_FileReadError));
    reader.readAsArrayBuffer(file);
  });
}

// ─── Apply mapping & normalize ────────────────────────────────────────────────

export function applyMapping(
  rows: Record<string, string>[],
  mapping: ColumnMapping
): Partial<ITask>[] {
  // A percent-formatted Excel cell ("50%") arrives via raw:true as the bare
  // fraction 0.5 with no literal '%' sign. Detect that once for the whole
  // mapped column (rather than per cell, where 0.5 could mean "half a
  // percent") and scale up if every non-empty, non-'%'-suffixed value is a
  // fraction in (0, 1].
  const percentCol = Object.keys(mapping).find(k => mapping[k] === 'percentComplete');
  let percentScale = 1;
  if (percentCol) {
    const values = rows
      .map(r => (r[percentCol] ?? '').trim())
      .filter(v => v !== '' && v.indexOf('%') === -1)
      .map(v => parseFloat(v))
      .filter(v => !isNaN(v));
    if (values.length > 0 && values.every(v => v > 0 && v <= 1)) {
      percentScale = 100;
    }
  }

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
            if (!isNaN(n)) task.percentComplete = Math.min(100, Math.max(0, n * percentScale));
            break;
          }
          case 'phase': task.phase = raw; break;
          case 'description': task.description = raw; break;
          case 'notes': task.notes = raw; break;
          case 'isMilestone': task.isMilestone = raw ? normalizeBoolean(raw) : false; break;
          case 'dependencies':
            // Not resolved here — SharePoint IDs don't exist until after
            // create. resolveDependencies() does a post-import pass over the
            // raw rows once all tasks have been created.
            break;
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

// Graph pages results regardless of $top (e.g. Planner tasks page at ~400
// per response); follow @odata.nextLink until exhausted so large collections
// aren't silently truncated.
async function fetchAllPages(client: IGraphClient, first: Promise<{ value?: any[]; '@odata.nextLink'?: string }>): Promise<any[]> {
  const results: any[] = [];
  let resp = await first;
  results.push(...(resp.value || []));
  let nextLink = resp['@odata.nextLink'];
  while (nextLink) {
    resp = await client.api(nextLink).get();
    results.push(...(resp.value || []));
    nextLink = resp['@odata.nextLink'];
  }
  return results;
}

export async function fetchPlannerPlans(context: WebPartContext): Promise<IPlannerPlan[]> {
  const graph = await getGraphClient(context);

  // Get the user's M365 groups
  let groups: any[] = [];
  try {
    groups = await fetchAllPages(
      graph,
      graph.api('/me/memberOf/microsoft.graph.group').select('id,displayName,groupTypes').top(50).get()
    );
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

  // Fetch tasks and buckets in parallel, paging both to completion.
  const [plannerTasks, buckets] = await Promise.all([
    fetchAllPages(graph, graph.api(`/planner/plans/${planId}/tasks`).top(500).get()),
    fetchAllPages(graph, graph.api(`/planner/plans/${planId}/buckets`).get()),
  ]);

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

// ─── Row filtering ────────────────────────────────────────────────────────────

// The same title-based filter applyMapping() uses to drop blank rows, exposed
// so callers can keep a raw-row array in lockstep with the filtered task
// array (needed to resolve dependencies positionally — see resolveDependencies).
export function filterMappedRows(rows: Record<string, string>[], mapping: ColumnMapping): Record<string, string>[] {
  const titleCol = Object.keys(mapping).find(k => mapping[k] === 'title');
  if (!titleCol) return [];
  return rows.filter(r => !!(r[titleCol] ?? '').trim());
}

// ─── Batch import ─────────────────────────────────────────────────────────────

export async function batchImport(
  spService: SharePointService,
  listName: string,
  tasks: Partial<ITask>[],
  sortOrderBase = 0,
  onProgress?: (done: number, total: number) => void
): Promise<IBatchImportResult> {
  const total = tasks.length;
  const prepared = tasks.map((t, i) => ({
    ...t,
    status: t.status || 'Not Started',
    priority: t.priority || 'Medium',
    percentComplete: t.percentComplete ?? 0,
    sortOrder: sortOrderBase + i,
  }));

  const CHUNK_SIZE = 50;
  const createdIds: Array<number | null> = new Array(total).fill(null);
  const errors: string[] = [];
  let succeeded = 0;
  let failed = 0;

  for (let start = 0; start < total; start += CHUNK_SIZE) {
    const chunk = prepared.slice(start, start + CHUNK_SIZE);
    const chunkResults = await spService.createTasksBatch(listName, chunk);
    chunkResults.forEach((r, i) => {
      const idx = start + i;
      if (r.id !== null) {
        createdIds[idx] = r.id;
        succeeded++;
      } else {
        failed++;
        errors.push(formatString(strings.ImportService_RowError, { rowNum: idx + 1, title: chunk[i].title, error: r.error || 'Unknown error' }));
      }
    });
    if (onProgress) onProgress(Math.min(start + CHUNK_SIZE, total), total);
  }

  return { succeeded, failed, errors, createdIds };
}

// ─── Post-import dependency resolution ───────────────────────────────────────
//
// Dependencies stored in Excel as task names (e.g. "Design Review, UX Wireframes")
// or MS Project-style row numbers can't be converted to SharePoint IDs at mapping
// time because the IDs don't exist yet. Call this after batchImport() with the
// same (filtered, aligned) rows and the createdIds it returned, to do a second
// pass wiring up each task's Dependencies field.
//
// `rows` MUST be the array returned by filterMappedRows() for the same mapping
// used to build the tasks passed to batchImport() — its order must match
// createdIds exactly, since "row number" dependencies and title lookups are
// both resolved positionally against createdIds rather than by re-fetching
// and re-matching titles from SharePoint (which breaks on duplicate titles).

export async function resolveDependencies(
  spService: SharePointService,
  listName: string,
  rows: Record<string, string>[],
  mapping: ColumnMapping,
  createdIds: Array<number | null>
): Promise<IResolveDependenciesResult> {
  const warnings: string[] = [];
  const depCol = Object.keys(mapping).find(k => mapping[k] === 'dependencies');
  const titleCol = Object.keys(mapping).find(k => mapping[k] === 'title');
  if (!depCol || !titleCol) return { resolved: 0, warnings };

  // 1-based row-number → SP id, so MS Project-style numeric predecessor
  // columns ("3", "5") resolve to the task actually created for that row.
  const rowToId = new Map<number, number>();
  rows.forEach((_row, idx) => {
    const id = createdIds[idx];
    if (id !== null && id !== undefined) rowToId.set(idx + 1, id);
  });

  // Case-insensitive title → SP id, tracking titles that appear more than
  // once so dependencies referencing them are skipped (with a warning)
  // rather than silently wired to whichever same-titled task happened last.
  const titleToId = new Map<string, number>();
  const ambiguousTitles = new Set<string>();
  rows.forEach((row, idx) => {
    const id = createdIds[idx];
    if (id === null || id === undefined) return;
    const title = (row[titleCol] ?? '').toLowerCase().trim();
    if (!title) return;
    if (titleToId.has(title)) {
      ambiguousTitles.add(title);
    } else {
      titleToId.set(title, id);
    }
  });

  const updates: Array<{ taskId: number; deps: number[] }> = [];

  rows.forEach((row, idx) => {
    const taskId = createdIds[idx];
    const rawDeps = (row[depCol] ?? '').trim();
    if (taskId === null || taskId === undefined || !rawDeps) return;

    const depIds: number[] = [];
    rawDeps.split(',').forEach(part => {
      const name = part.trim();
      if (!name) return;

      // Pure integer → treat as 1-based row number (MS Project style)
      const rowNum = parseInt(name, 10);
      if (!isNaN(rowNum) && String(rowNum) === name && rowNum > 0) {
        const id = rowToId.get(rowNum);
        if (id !== undefined) depIds.push(id);
        return;
      }

      // Otherwise match by task title (case-insensitive)
      const key = name.toLowerCase();
      if (ambiguousTitles.has(key)) {
        warnings.push(formatString(strings.ImportService_AmbiguousDependency, { name }));
        return;
      }
      const id = titleToId.get(key);
      if (id !== undefined) depIds.push(id);
    });

    if (depIds.length > 0) updates.push({ taskId, deps: depIds });
  });

  if (updates.length === 0) return { resolved: 0, warnings };

  // Promise.allSettled isn't available at this project's target lib — a
  // per-item catch gives the same "don't let one failure abort the batch"
  // behavior without it.
  interface IUpdateOutcome { error: string | null; }
  const CHUNK_SIZE = 10;
  let resolved = 0;
  for (let start = 0; start < updates.length; start += CHUNK_SIZE) {
    const chunk = updates.slice(start, start + CHUNK_SIZE);
    const results: IUpdateOutcome[] = await Promise.all(
      chunk.map(({ taskId, deps }): Promise<IUpdateOutcome> =>
        spService.updateTask(listName, taskId, { dependencies: deps })
          .then((): IUpdateOutcome => ({ error: null }))
          .catch((e: unknown): IUpdateOutcome => ({ error: e instanceof Error ? e.message : String(e) }))
      )
    );
    results.forEach((r, i) => {
      if (r.error === null) {
        resolved++;
      } else {
        warnings.push(formatString(strings.ImportService_LinkDependencyError, { taskId: chunk[i].taskId, error: r.error }));
      }
    });
  }

  return { resolved, warnings };
}
