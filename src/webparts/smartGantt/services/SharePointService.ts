import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/fields';
import '@pnp/sp/views';
import '@pnp/sp/batching';
import '@pnp/sp/profiles';
import * as strings from 'SmartGanttWebPartStrings';
import { IDependencyLink, IProject, ITask, IProjectTaskStats, ProjectStatus, TaskPriority, TaskStatus } from '../models';
import { computeTaskHealth, computeProjectHealth } from '../utils/healthUtils';
import { toDateOnly } from '../utils/dateUtils';
import { parseDependencies, serializeDependencies } from '../utils/dependencyUtils';
import { withRetry, runLimited } from '../utils/retryUtils';
import { formatString } from '../components/localeUtils';

const PROJECTS_LIST = 'SmartGantt_Projects';

// Schedule dates are calendar days. The columns are DateTime fields that
// SharePoint displays in the site's time zone, so UTC midnight showed as the
// previous day on sites west of UTC. We store 10:00 UTC instead, which is on
// the same calendar day for every site zone from UTC-10 (Hawaii) to UTC+13
// (NZ daylight time, Tonga); no single instant can also cover UTC-11 and +14.
// Reading always goes through toDateOnly() (utils/dateUtils), which takes the
// UTC calendar date, so the app shows the stored day in every viewer zone and
// older values written as UTC midnight still read as the same day.
// PnPjs serializes Date objects for Edm.DateTime fields; raw ISO strings can
// fail in some tenants.
function toSPDate(s: string | undefined | null): Date | null {
  const dateOnly = toDateOnly(s);
  if (!dateOnly) return null;
  const [y, m, d] = dateOnly.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d, 10, 0, 0));
}

// PnPjs surfaces HTTP failures as HttpRequestError with a status property;
// only a 404 means "the list/field genuinely doesn't exist". Anything else
// (403 for a read-only user, throttling, network) must not trigger creation.
function isNotFoundError(e: unknown): boolean {
  const status = (e as { status?: number })?.status;
  if (status === 404) return true;
  if (status !== undefined) return false;
  const msg = e instanceof Error ? e.message : String(e);
  return msg.indexOf('404') !== -1 || /does not exist/i.test(msg);
}

// The free-name probe in _buildListName and the actual sp.web.lists.add()
// call aren't atomic, so two concurrent creates can both see a name as free.
// SharePoint's duplicate-title error is only distinguishable by message text.
function isDuplicateListNameError(e: unknown): boolean {
  const msg = e instanceof Error ? e.message : String(e);
  return /already exists/i.test(msg);
}

// Thrown by SharePoint when an If-Match eTag no longer matches — i.e. someone
// else saved this item since it was read.
function isPreconditionFailedError(e: unknown): boolean {
  const status = (e as { status?: number })?.status;
  if (status === 412) return true;
  const msg = e instanceof Error ? e.message : String(e);
  return /412/.test(msg) || /precondition/i.test(msg);
}

type ConflictError = Error & { status: number; isConflict: true };

function makeConflictError(message: string): ConflictError {
  const err = new Error(message) as ConflictError;
  err.status = 412;
  err.isConflict = true;
  return err;
}

/** True for the error updateTask/updateProject throw when the item changed since it was read (HTTP 412/409). */
export function isConflictError(e: unknown): boolean {
  const err = e as { isConflict?: boolean; status?: number } | null;
  return !!err && (err.isConflict === true || err.status === 412 || err.status === 409);
}

// The Dependencies field is a plain Text column capped at 500 characters
// (see field def below); a task with too many dependencies would otherwise
// fail the save with a raw SharePoint error deep inside a PnPjs call.
const DEPENDENCIES_MAX_LENGTH = 500;
function joinDependencies(ids: number[] | undefined, links?: Record<number, IDependencyLink>): string {
  const joined = serializeDependencies(ids || [], links);
  if (joined.length > DEPENDENCIES_MAX_LENGTH) {
    throw new Error(strings.Svc_TooManyDependencies);
  }
  return joined;
}

const STATS_TTL_MS = 60000;
const STATS_CONCURRENCY = 4;

export class SharePointService {
  private sp: SPFI;
  private projectsListEnsured = false;
  // IsArchived migration is attempted at most once per service instance, so a
  // read-only user doesn't repeat a doomed field-add (and log) on every load.
  private isArchivedChecked = false;
  // listName -> whether BaselineStart/BaselineDue exist (or could be added).
  private baselineSupport = new Map<string, boolean>();
  private statsCache = new Map<string, { at: number; stats: IProjectTaskStats }>();

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ─── Projects ────────────────────────────────────────────────────────────

  async ensureProjectsList(): Promise<void> {
    if (this.projectsListEnsured) return;
    try {
      await this.sp.web.lists.getByTitle(PROJECTS_LIST)();
      // List exists — add IsArchived if missing (migration for deployments before v1.2)
      await this._ensureIsArchivedField();
    } catch (e) {
      if (!isNotFoundError(e)) throw e;
      await this.sp.web.lists.add(PROJECTS_LIST, 'Smart Gantt — project registry', 100, false);
      const list = this.sp.web.lists.getByTitle(PROJECTS_LIST);
      await list.fields.addText('ProjectListName', { MaxLength: 255 });
      await list.fields.addMultilineText('ProjectDescription', { NumberOfLines: 4 });
      await list.fields.addText('ProjectColor', { MaxLength: 20 });
      await list.fields.addDateTime('ProjectStartDate');
      await list.fields.addDateTime('ProjectDueDate');
      await list.fields.addChoice('ProjectStatus', {
        Choices: ['Planning', 'Active', 'On Hold', 'Completed', 'Cancelled'],
      });
      await list.fields.addText('ProjectManager', { MaxLength: 255 });
      await list.fields.addText('ProjectManagerEmail', { MaxLength: 255 });
      await list.fields.add('IsArchived', 8, {});
      await this._setupMetaListView();
    }
    this.projectsListEnsured = true;
  }

  private async _ensureIsArchivedField(): Promise<void> {
    if (this.isArchivedChecked) return;
    const list = this.sp.web.lists.getByTitle(PROJECTS_LIST);
    try {
      await withRetry(() => list.fields.getByInternalNameOrTitle('IsArchived')());
      this.isArchivedChecked = true;
    } catch (e) {
      // Only a genuine "field not found" means the migration is needed; a 403,
      // throttling or network error says nothing about the field, so it must
      // not trigger a field-add. Either way, don't probe again this session.
      this.isArchivedChecked = true;
      if (!isNotFoundError(e)) {
        console.warn('[SmartGantt] IsArchived check failed (skipping migration):', e);
        return;
      }
      try {
        await list.fields.add('IsArchived', 8, {});
      } catch (addErr) {
        // Read-only users can't add fields; archiving simply won't be available.
        console.warn('[SmartGantt] IsArchived migration skipped:', addErr);
      }
    }
  }

  async getProjects(): Promise<IProject[]> {
    await this.ensureProjectsList();
    const items = await withRetry(() => this.sp.web.lists
      .getByTitle(PROJECTS_LIST)
      .items.select(
        'Id', 'Title', 'ProjectListName', 'ProjectDescription', 'ProjectColor',
        'ProjectStartDate', 'ProjectDueDate', 'ProjectStatus', 'ProjectManager',
        'ProjectManagerEmail', 'Created', 'IsArchived'
      )
      .getAll());

    return items
      .map(item => ({
        id: item.Id,
        title: item.Title,
        listName: item.ProjectListName || '',
        description: item.ProjectDescription || '',
        color: item.ProjectColor || '#0078D4',
        startDate: toDateOnly(item.ProjectStartDate),
        dueDate: toDateOnly(item.ProjectDueDate),
        status: (item.ProjectStatus || 'Active') as ProjectStatus,
        projectManager: item.ProjectManager || '',
        projectManagerEmail: item.ProjectManagerEmail || '',
        created: item.Created,
        isArchived: item.IsArchived === true,
        // OData control metadata is present on every item regardless of the
        // $select projection above, so no explicit select entry is needed.
        etag: item['odata.etag'] as string | undefined,
      }))
      .sort((a, b) => a.title.localeCompare(b.title));
  }

  async createProject(data: {
    title: string;
    description: string;
    color: string;
    startDate: string;
    dueDate: string;
    status: ProjectStatus;
  }): Promise<IProject> {
    await this.ensureProjectsList();

    const sanitized = data.title.replace(/[^a-zA-Z0-9]/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '') || 'Project';
    let listName = await this._buildListName(sanitized);

    try {
      await this._createProjectList(listName);
    } catch (e) {
      // The free-name probe and this create call aren't atomic — retry once
      // with a name a concurrent create couldn't have raced on.
      if (!isDuplicateListNameError(e)) throw e;
      listName = `${listName}_${Date.now().toString(36).toUpperCase()}`;
      await this._createProjectList(listName);
    }

    let result;
    try {
      result = await this.sp.web.lists.getByTitle(PROJECTS_LIST).items.add({
        Title: data.title,
        ProjectListName: listName,
        ProjectDescription: data.description,
        ProjectColor: data.color,
        ProjectStartDate: toSPDate(data.startDate),
        ProjectDueDate: toSPDate(data.dueDate),
        ProjectStatus: data.status,
      });
    } catch (e) {
      // Registry entry failed — recycle the orphan task list so a retry is clean.
      try { await this.sp.web.lists.getByTitle(listName).recycle(); } catch { /* best effort */ }
      throw e;
    }

    return {
      id: result.data.Id,
      title: data.title,
      listName,
      description: data.description,
      color: data.color,
      startDate: toDateOnly(data.startDate),
      dueDate: toDateOnly(data.dueDate),
      status: data.status,
      projectManager: '',
      projectManagerEmail: '',
      created: new Date().toISOString(),
    };
  }

  async updateProject(id: number, data: Partial<IProject>): Promise<void> {
    const updates: Record<string, unknown> = {};
    if (data.title !== undefined) updates.Title = data.title;
    if (data.description !== undefined) updates.ProjectDescription = data.description;
    if (data.color !== undefined) updates.ProjectColor = data.color;
    if (data.startDate !== undefined) updates.ProjectStartDate = toSPDate(data.startDate);
    if (data.dueDate !== undefined) updates.ProjectDueDate = toSPDate(data.dueDate);
    if (data.status !== undefined) updates.ProjectStatus = data.status;
    if (data.isArchived !== undefined) updates.IsArchived = data.isArchived;
    if (data.projectManager !== undefined) updates.ProjectManager = data.projectManager;
    if (data.projectManagerEmail !== undefined) updates.ProjectManagerEmail = data.projectManagerEmail;
    try {
      await this.sp.web.lists.getByTitle(PROJECTS_LIST).items.getById(id).update(updates, data.etag || '*');
      // Project dates feed the health figures in cached stats.
      this.invalidateStats();
    } catch (e) {
      if (isPreconditionFailedError(e)) {
        throw makeConflictError(strings.Svc_ProjectChangedByOthers);
      }
      throw e;
    }
  }

  async archiveProject(id: number): Promise<void> {
    await this.sp.web.lists.getByTitle(PROJECTS_LIST).items.getById(id).update({ IsArchived: true });
  }

  async unarchiveProject(id: number): Promise<void> {
    await this.sp.web.lists.getByTitle(PROJECTS_LIST).items.getById(id).update({ IsArchived: false });
  }

  async deleteProject(id: number, listName: string): Promise<void> {
    // Recycle the task list FIRST. If that fails (e.g. 403) the registry entry
    // is still intact, so the project stays visible and the delete can be
    // retried; the reverse order would leave an orphaned, unreachable list.
    try {
      await this.sp.web.lists.getByTitle(listName).recycle();
    } catch (e) {
      // A missing list is fine (already gone). Anything else must surface.
      if (!isNotFoundError(e)) throw e;
    }
    await this.sp.web.lists.getByTitle(PROJECTS_LIST).items.getById(id).recycle();
    this.invalidateStats(listName);
    this.baselineSupport.delete(listName);
  }

  private async _createProjectList(listName: string): Promise<void> {
    await this.sp.web.lists.add(listName, 'Task list for Smart Gantt project', 100, false);

    try {
      // Brief pause so SharePoint fully provisions the list before we add fields.
      await new Promise<void>(resolve => setTimeout(resolve, 1500));

      // IsMilestone uses FieldTypeKind 8 (Boolean) via the generic add() because
      // addBoolean() in PnPjs 3.x passes the wrong SP type and triggers a 400.
      // All field adds go through a single REST batch — one round trip instead
      // of seventeen. Any failed field fails the whole create (the catch below
      // deletes the half-built list) so a project never ends up without columns.
      const [batchedSP, execute] = this.sp.batched();
      const list = batchedSP.web.lists.getByTitle(listName);
      const pending: Array<{ name: string; promise: Promise<unknown>; fallback?: () => Promise<unknown>; optional?: boolean }> = [];
      const queue = (
        name: string,
        promise: Promise<unknown>,
        opts: { fallback?: () => Promise<unknown>; optional?: boolean } = {},
      ): void => { pending.push({ name, promise, ...opts }); };
      const plainList = this.sp.web.lists.getByTitle(listName);

      queue('TaskDescription', list.fields.addMultilineText('TaskDescription'));
      queue('StartDate', list.fields.addDateTime('StartDate'));
      queue('DueDate', list.fields.addDateTime('DueDate'));
      queue('Status', list.fields.addChoice('Status', {
        Choices: ['Not Started', 'In Progress', 'Completed', 'On Hold', 'Cancelled'],
      }));
      queue('Priority', list.fields.addChoice('Priority', {
        Choices: ['Critical', 'High', 'Medium', 'Low'],
      }));
      queue('PercentComplete', list.fields.addNumber('PercentComplete'));
      // Indexing is a nice-to-have (used by the parent-task lookup); if the
      // index can't be created, a plain number column still works.
      queue('ParentTaskId', list.fields.addNumber('ParentTaskId', { Indexed: true }), {
        fallback: () => plainList.fields.addNumber('ParentTaskId'),
      });
      // Single-line text columns are capped at 255 characters by SharePoint, so
      // a longer MaxLength is rejected; fall back to a multi-line text column,
      // which has no practical limit and reads back as the same plain string.
      queue('Dependencies', list.fields.addText('Dependencies', { MaxLength: Math.min(DEPENDENCIES_MAX_LENGTH, 255) }), {
        fallback: () => plainList.fields.addMultilineText('Dependencies'),
      });
      queue('Notes', list.fields.addMultilineText('Notes'));
      queue('TaskColor', list.fields.addText('TaskColor', { MaxLength: 20 }));
      queue('SortOrder', list.fields.addNumber('SortOrder'));
      queue('IsMilestone', list.fields.add('IsMilestone', 8));
      queue('Phase', list.fields.addText('Phase', { MaxLength: 100 }));
      queue('AssignedToName', list.fields.addText('AssignedToName', { MaxLength: 255 }));
      queue('AssignedToEmail', list.fields.addText('AssignedToEmail', { MaxLength: 255 }));
      // Baselines are optional: _ensureBaselineFields adds them lazily later.
      queue('BaselineStart', list.fields.addDateTime('BaselineStart'), { optional: true });
      queue('BaselineDue', list.fields.addDateTime('BaselineDue'), { optional: true });

      const outcomes = pending.map(p => p.promise.then(() => null, (e: unknown) => e));
      await execute();
      const results = await Promise.all(outcomes);

      const failures: Array<{ name: string; error: unknown }> = [];
      for (let i = 0; i < pending.length; i++) {
        if (results[i] === null) continue;
        const p = pending[i];
        if (p.optional) {
          console.warn(`[SmartGantt] Optional column ${p.name} was not created:`, results[i]);
          continue;
        }
        if (p.fallback) {
          try {
            await p.fallback();
            console.warn(`[SmartGantt] Column ${p.name} created with fallback settings after:`, results[i]);
            continue;
          } catch (e) {
            failures.push({ name: p.name, error: e });
            continue;
          }
        }
        failures.push({ name: p.name, error: results[i] });
      }
      if (failures.length > 0) {
        const detail = failures
          .map(f => `${f.name}: ${f.error instanceof Error ? f.error.message : String(f.error)}`)
          .join(' | ');
        throw new Error(formatString(strings.Svc_ColumnCreateFailed, {
          count: failures.length,
          detail,
        }));
      }
      this.baselineSupport.set(listName, !pending.some((p, i) => p.optional && results[i] !== null));

      await this._setupTaskListViews(listName);
    } catch (e) {
      // Anything past this point failing would otherwise leave the list
      // behind with no registry entry pointing to it and no cleanup —
      // createProject()'s own try/catch only covers the later registry-add
      // step, not this one.
      try { await this.sp.web.lists.getByTitle(listName).recycle(); } catch { /* best effort */ }
      throw e;
    }
  }

  // ─── Tasks ────────────────────────────────────────────────────────────────

  /**
   * Whether the task list has (or could be given) the BaselineStart/BaselineDue
   * columns. Lists created before baselines existed get them added lazily; a
   * user without permission to add fields simply gets no baseline support for
   * that list (cached, so we don't retry on every read).
   */
  private async _ensureBaselineFields(listName: string): Promise<boolean> {
    const cached = this.baselineSupport.get(listName);
    if (cached !== undefined) return cached;
    let supported = false;
    try {
      const list = this.sp.web.lists.getByTitle(listName);
      const existing = await withRetry(() => list.fields
        .select('InternalName')
        .filter("InternalName eq 'BaselineStart' or InternalName eq 'BaselineDue'")());
      const have = new Set<string>((existing as Array<{ InternalName: string }>).map(f => f.InternalName));
      try {
        if (!have.has('BaselineStart')) await list.fields.addDateTime('BaselineStart');
        if (!have.has('BaselineDue')) await list.fields.addDateTime('BaselineDue');
        supported = true;
      } catch (e) {
        // Read-only / limited users can't add fields; baselines just aren't available.
        console.warn('[SmartGantt] Baseline field migration skipped:', e);
      }
    } catch (e) {
      // Don't cache a transient failure — try again on the next read.
      console.warn('[SmartGantt] Baseline field check failed:', e);
      return false;
    }
    this.baselineSupport.set(listName, supported);
    return supported;
  }

  async getProjectTasks(listName: string): Promise<ITask[]> {
    const baselines = await this._ensureBaselineFields(listName);
    const fields = [
      'Id', 'Title', 'TaskDescription', 'StartDate', 'DueDate', 'Status', 'Priority',
      'AssignedToName', 'AssignedToEmail', 'PercentComplete', 'ParentTaskId', 'Dependencies',
      'Notes', 'TaskColor', 'SortOrder', 'IsMilestone', 'Phase', 'Created', 'Modified',
    ];
    if (baselines) fields.push('BaselineStart', 'BaselineDue');
    const items = await withRetry(() => this.sp.web.lists
      .getByTitle(listName)
      .items.select(...fields)
      .getAll());

    return items
      .map(item => {
        const parsed = parseDependencies(item.Dependencies || '');
        const task: ITask = {
          id: item.Id,
          title: item.Title,
          description: item.TaskDescription || '',
          startDate: toDateOnly(item.StartDate),
          dueDate: toDateOnly(item.DueDate),
          status: (item.Status || 'Not Started') as TaskStatus,
          priority: (item.Priority || 'Medium') as TaskPriority,
          assignedTo: item.AssignedToName || '',
          assignedToEmail: item.AssignedToEmail || '',
          percentComplete: item.PercentComplete || 0,
          parentTaskId: item.ParentTaskId || null,
          dependencies: parsed.ids,
          dependencyLinks: parsed.links,
          notes: item.Notes || '',
          color: item.TaskColor || '',
          sortOrder: item.SortOrder || 0,
          isMilestone: item.IsMilestone === true || item.IsMilestone === 1,
          phase: item.Phase || '',
          created: item.Created,
          modified: item.Modified,
          // Present on every item regardless of the $select projection.
          etag: item['odata.etag'] as string | undefined,
        };
        if (baselines) {
          const bs = toDateOnly(item.BaselineStart);
          const bd = toDateOnly(item.BaselineDue);
          if (bs) task.baselineStart = bs;
          if (bd) task.baselineDue = bd;
        }
        return task;
      })
      .sort((a, b) => a.sortOrder - b.sortOrder || a.id - b.id);
  }

  /** Drop cached portfolio stats for one task list, or all of them. */
  invalidateStats(listName?: string): void {
    if (listName) this.statsCache.delete(listName); else this.statsCache.clear();
  }

  async getProjectTaskStats(project: IProject, force = false): Promise<IProjectTaskStats> {
    const cached = this.statsCache.get(project.listName);
    if (!force && cached && Date.now() - cached.at < STATS_TTL_MS) return cached.stats;
    const stats = await this._loadProjectTaskStats(project);
    // Failures are not cached, so a transient error can be retried immediately.
    if (!stats.statsError) this.statsCache.set(project.listName, { at: Date.now(), stats });
    return stats;
  }

  private async _loadProjectTaskStats(project: IProject): Promise<IProjectTaskStats> {
    const empty: IProjectTaskStats = {
      listName: project.listName,
      totalTasks: 0,
      byStatus: { 'Not Started': 0, 'In Progress': 0, 'Completed': 0, 'On Hold': 0, 'Cancelled': 0 },
      overallPct: 0,
      health: 'on-track',
      overdueCount: 0,
      atRiskCount: 0,
      inProgressCount: 0,
      completedCount: 0,
      milestoneCount: 0,
      earliestStart: project.startDate || '',
      latestDue: project.dueDate || '',
    };

    try {
      // Only the columns the stats need, in large pages (fewer round trips).
      const items = await withRetry(() => this.sp.web.lists
        .getByTitle(project.listName)
        .items.select('Status', 'Priority', 'PercentComplete', 'StartDate', 'DueDate', 'IsMilestone')
        .getAll(2000));

      if (items.length === 0) return empty;

      const tasks: ITask[] = items.map(item => ({
        id: 0,
        title: '',
        description: '',
        startDate: toDateOnly(item.StartDate),
        dueDate: toDateOnly(item.DueDate),
        status: (item.Status || 'Not Started') as TaskStatus,
        priority: (item.Priority || 'Medium') as TaskPriority,
        assignedTo: '', assignedToEmail: '',
        percentComplete: item.PercentComplete || 0,
        parentTaskId: null, dependencies: [], notes: '', color: '',
        sortOrder: 0,
        isMilestone: item.IsMilestone === true || item.IsMilestone === 1,
        phase: '', created: '', modified: '',
      }));

      const today = new Date();
      const byStatus = { 'Not Started': 0, 'In Progress': 0, 'Completed': 0, 'On Hold': 0, 'Cancelled': 0 } as Record<TaskStatus, number>;
      let totalPct = 0;
      let overdueCount = 0;
      let atRiskCount = 0;
      let milestoneCount = 0;
      let earliestStart = '';
      let latestDue = '';

      for (const task of tasks) {
        byStatus[task.status] = (byStatus[task.status] || 0) + 1;
        totalPct += task.percentComplete;
        if (task.isMilestone) milestoneCount++;
        if (task.startDate && (!earliestStart || task.startDate < earliestStart)) earliestStart = task.startDate;
        if (task.dueDate && (!latestDue || task.dueDate > latestDue)) latestDue = task.dueDate;
        const h = computeTaskHealth(task, today);
        if (h === 'overdue') overdueCount++;
        if (h === 'at-risk') atRiskCount++;
      }

      return {
        listName: project.listName,
        totalTasks: tasks.length,
        byStatus,
        overallPct: Math.round(totalPct / tasks.length),
        health: computeProjectHealth(tasks, project, today),
        overdueCount,
        atRiskCount,
        inProgressCount: byStatus['In Progress'],
        completedCount: byStatus['Completed'],
        milestoneCount,
        earliestStart: earliestStart || project.startDate || '',
        latestDue: latestDue || project.dueDate || '',
      };
    } catch (e) {
      // A genuinely missing list (project registry entry outlived its task
      // list) really is an empty/on-track project. Anything else — 403,
      // throttling, network — must not be reported as healthy; the caller
      // should render it as unavailable instead.
      if (isNotFoundError(e)) return empty;
      return { ...empty, statsError: true };
    }
  }

  /** Stats for every project, STATS_CONCURRENCY lists at a time; results are cached (see invalidateStats). */
  async getAllProjectStats(projects: IProject[], force = false): Promise<Map<number, IProjectTaskStats>> {
    const map = new Map<number, IProjectTaskStats>();
    const results = await runLimited(projects, STATS_CONCURRENCY,
      p => this.getProjectTaskStats(p, force).then(stats => ({ id: p.id, stats })));
    results.forEach(r => map.set(r.id, r.stats));
    return map;
  }

  private _taskCreateFields(task: Partial<ITask>): Record<string, unknown> {
    return {
      Title: task.title || 'New Task',
      TaskDescription: task.description || '',
      StartDate: toSPDate(task.startDate),
      DueDate: toSPDate(task.dueDate),
      Status: task.status || 'Not Started',
      Priority: task.priority || 'Medium',
      AssignedToName: task.assignedTo || '',
      AssignedToEmail: task.assignedToEmail || '',
      PercentComplete: task.percentComplete || 0,
      ParentTaskId: task.parentTaskId || 0,
      Dependencies: joinDependencies(task.dependencies, task.dependencyLinks),
      Notes: task.notes || '',
      TaskColor: task.color || '',
      SortOrder: task.sortOrder || 0,
      IsMilestone: task.isMilestone || false,
      Phase: task.phase || '',
    };
  }

  async createTask(listName: string, task: Partial<ITask>): Promise<ITask> {
    const result = await this.sp.web.lists.getByTitle(listName).items.add(this._taskCreateFields(task));
    this.invalidateStats(listName);

    // items.add() returns the created row's data; its concurrency token (when
    // present) lets the caller update the new task straight away.
    const created = result.data as Record<string, unknown>;
    return {
      id: result.data.Id,
      title: task.title || 'New Task',
      description: task.description || '',
      startDate: toDateOnly(task.startDate),
      dueDate: toDateOnly(task.dueDate),
      status: (task.status || 'Not Started') as TaskStatus,
      priority: (task.priority || 'Medium') as TaskPriority,
      assignedTo: task.assignedTo || '',
      assignedToEmail: task.assignedToEmail || '',
      percentComplete: task.percentComplete || 0,
      parentTaskId: task.parentTaskId || null,
      dependencies: task.dependencies || [],
      dependencyLinks: task.dependencyLinks,
      notes: task.notes || '',
      color: task.color || '',
      sortOrder: task.sortOrder || 0,
      isMilestone: task.isMilestone || false,
      phase: task.phase || '',
      created: new Date().toISOString(),
      modified: new Date().toISOString(),
      etag: (created['odata.etag'] as string | undefined) || undefined,
    };
  }

  /**
   * Create many tasks in chunked REST batches (instead of one round trip per
   * task) and report a per-row result in input order, so callers can align a
   * created-id array back against their original rows (e.g. to resolve
   * dependencies positionally after a bulk import).
   */
  async createTasksBatch(
    listName: string,
    tasks: Partial<ITask>[],
    chunkSize = 50
  ): Promise<Array<{ id: number | null; error?: string }>> {
    const results: Array<{ id: number | null; error?: string }> = tasks.map(() => ({ id: null }));

    for (let start = 0; start < tasks.length; start += chunkSize) {
      const chunk = tasks.slice(start, start + chunkSize);
      // A whole-batch throttle (429/503/504) means nothing in it was applied,
      // so re-running the same chunk can't create duplicates. Per-row errors
      // are reported in `results` instead.
      await withRetry(async () => {
        const [batchedSP, execute] = this.sp.batched();
        const list = batchedSP.web.lists.getByTitle(listName);
        chunk.forEach((task, i) => {
          const idx = start + i;
          list.items.add(this._taskCreateFields(task))
            .then(r => { results[idx] = { id: r.data.Id }; })
            .catch((e: unknown) => {
              results[idx] = { id: null, error: e instanceof Error ? e.message : 'Unknown error' };
            });
        });
        await execute();
      });
    }

    this.invalidateStats(listName);
    return results;
  }

  /**
   * Update a task. Sends `updates.etag` as If-Match (falls back to '*'), and
   * resolves with the item's NEW etag so the caller can keep editing it. On a
   * 412 the thrown error is a conflict error (see isConflictError) with a
   * localized "changed by someone else" message.
   */
  async updateTask(listName: string, id: number, updates: Partial<ITask> & { etag?: string }): Promise<string | undefined> {
    const data: Record<string, unknown> = {};
    if (updates.title !== undefined) data.Title = updates.title;
    if (updates.description !== undefined) data.TaskDescription = updates.description;
    if (updates.startDate !== undefined) data.StartDate = toSPDate(updates.startDate);
    if (updates.dueDate !== undefined) data.DueDate = toSPDate(updates.dueDate);
    if (updates.status !== undefined) data.Status = updates.status;
    if (updates.priority !== undefined) data.Priority = updates.priority;
    if (updates.assignedTo !== undefined) data.AssignedToName = updates.assignedTo;
    if (updates.assignedToEmail !== undefined) data.AssignedToEmail = updates.assignedToEmail;
    if (updates.percentComplete !== undefined) data.PercentComplete = updates.percentComplete;
    if (updates.parentTaskId !== undefined) data.ParentTaskId = updates.parentTaskId || 0;
    if (updates.dependencies !== undefined || updates.dependencyLinks !== undefined) {
      let ids = updates.dependencies;
      let links = updates.dependencyLinks;
      if (ids === undefined) {
        // Links without an id list: the links' keys are the predecessors.
        ids = Object.keys(links || {}).map(Number);
      } else if (links === undefined) {
        // Ids changed but no link info passed: keep the link types the task
        // already has for predecessors that remain, instead of silently
        // resetting them to plain finish-to-start.
        const current = await this.sp.web.lists.getByTitle(listName).items.getById(id).select('Dependencies')();
        links = parseDependencies(current.Dependencies || '').links;
      }
      data.Dependencies = joinDependencies(ids, links);
    }
    if (updates.notes !== undefined) data.Notes = updates.notes;
    if (updates.color !== undefined) data.TaskColor = updates.color;
    if (updates.sortOrder !== undefined) data.SortOrder = updates.sortOrder;
    if (updates.isMilestone !== undefined) data.IsMilestone = updates.isMilestone;
    if (updates.phase !== undefined) data.Phase = updates.phase;
    if (updates.baselineStart !== undefined || updates.baselineDue !== undefined) {
      if (await this._ensureBaselineFields(listName)) {
        if (updates.baselineStart !== undefined) data.BaselineStart = toSPDate(updates.baselineStart);
        if (updates.baselineDue !== undefined) data.BaselineDue = toSPDate(updates.baselineDue);
      }
    }

    const item = this.sp.web.lists.getByTitle(listName).items.getById(id);
    let newEtag: string | undefined;
    try {
      const res = await withRetry(() => item.update(data, updates.etag || '*'));
      newEtag = (res.data as { etag?: string | null } | undefined)?.etag || undefined;
    } catch (e) {
      if (isPreconditionFailedError(e)) {
        throw makeConflictError(strings.Svc_TaskChangedByOthers);
      }
      throw e;
    }
    this.invalidateStats(listName);

    if (!newEtag) {
      // The MERGE response normally carries the ETag header; if a proxy
      // stripped it, read it back so the caller never keeps a stale token.
      try {
        const fresh = await item.select('Id')();
        newEtag = (fresh as Record<string, unknown>)['odata.etag'] as string | undefined;
      } catch { /* best effort — caller falls back to '*' */ }
    }
    return newEtag;
  }

  /**
   * Snapshot every dated task's current start/due into BaselineStart/BaselineDue
   * (batched, `*` etag — it overwrites regardless of concurrent edits). Task
   * etags change as a result, so callers should reload the tasks afterwards.
   */
  async captureBaseline(listName: string, tasks: ITask[]): Promise<void> {
    if (!(await this._ensureBaselineFields(listName))) {
      throw new Error(strings.Svc_BaselineUnavailable);
    }
    const CHUNK_SIZE = 50;
    const errors: unknown[] = [];
    for (let start = 0; start < tasks.length; start += CHUNK_SIZE) {
      const chunk = tasks.slice(start, start + CHUNK_SIZE);
      await withRetry(async () => {
        const [batchedSP, execute] = this.sp.batched();
        const list = batchedSP.web.lists.getByTitle(listName);
        chunk.forEach(t => {
          list.items.getById(t.id)
            .update({ BaselineStart: toSPDate(t.startDate), BaselineDue: toSPDate(t.dueDate) })
            .catch(e => errors.push(e));
        });
        await execute();
      });
    }
    this.invalidateStats(listName);
    if (errors.length > 0) {
      const first = errors[0];
      throw new Error(first instanceof Error ? first.message : String(first));
    }
  }

  /**
   * Directory search for a people picker: matches from the tenant's people
   * picker (any user the current user can resolve), by name or email. Never
   * throws — a failed or unauthorized search just yields no suggestions.
   */
  async searchPeople(query: string): Promise<Array<{ name: string; email: string }>> {
    const q = (query || '').trim();
    if (q.length < 2) return [];
    try {
      const entities = await withRetry(() => this.sp.profiles.clientPeoplePickerSearchUser({
        QueryString: q,
        MaximumEntitySuggestions: 10,
        AllowEmailAddresses: true,
        AllowMultipleEntities: false,
        AllUrlZones: false,
        // PrincipalType.User (1) and PrincipalSource.All (15); the PnP enums
        // are ambient declarations, so pass the numeric values.
        PrincipalType: 1 as never,
        PrincipalSource: 15 as never,
      }), { retries: 1 });
      const seen = new Set<string>();
      const people: Array<{ name: string; email: string }> = [];
      (entities || []).forEach(en => {
        const email = (en.EntityData && en.EntityData.Email) || '';
        const name = en.DisplayText || email;
        const key = (email || name).toLowerCase();
        if (!name || seen.has(key)) return;
        seen.add(key);
        people.push({ name, email });
      });
      return people;
    } catch {
      return [];
    }
  }

  async deleteTask(listName: string, id: number): Promise<void> {
    const list = this.sp.web.lists.getByTitle(listName);

    // Promote sub-tasks to top level first, so they don't become invisible
    // orphans (views only render sub-tasks under an existing parent).
    // getAll() (rather than a single top(500)) pages through every child,
    // and any promotion failure aborts the delete instead of proceeding to
    // recycle the parent — a partially-promoted set of children would
    // otherwise be silently orphaned under a since-recycled parent.
    try {
      const children = await withRetry(() => list.items.select('Id').filter(`ParentTaskId eq ${id}`).getAll());
      const CHUNK_SIZE = 100;
      for (let start = 0; start < children.length; start += CHUNK_SIZE) {
        const chunk = children.slice(start, start + CHUNK_SIZE);
        const errors: unknown[] = [];
        await withRetry(async () => {
          errors.length = 0;
          const [batchedSP, execute] = this.sp.batched();
          const batchedList = batchedSP.web.lists.getByTitle(listName);
          chunk.forEach(c => {
            batchedList.items.getById(c.Id).update({ ParentTaskId: 0 }).catch(e => errors.push(e));
          });
          await execute();
        });
        if (errors.length > 0) throw errors[0];
      }
    } catch (e) {
      throw new Error(
        formatString(strings.Svc_SubtaskReassignFailed, { detail: e instanceof Error ? e.message : 'unknown error' })
      );
    }

    // Recycle (not delete) so the task can be restored from the recycle bin.
    await list.items.getById(id).recycle();
    this.invalidateStats(listName);
  }

  // ─── List naming ──────────────────────────────────────────────────────────

  private async _buildListName(sanitized: string): Promise<string> {
    const base = `SmartGantt_${sanitized.substring(0, 50)}`;
    try {
      await this.sp.web.lists.getByTitle(base)();
      // Base name is taken — try _2, _3, …
      for (let n = 2; n <= 99; n++) {
        const candidate = `${base}_${n}`;
        try {
          await this.sp.web.lists.getByTitle(candidate)();
        } catch (e) {
          if (isNotFoundError(e)) return candidate; // free
          throw e; // 403/throttling on the probe — not a real "taken" signal
        }
      }
      return `${base}_${Date.now().toString(36).toUpperCase()}`; // extremely unlikely fallback
    } catch (e) {
      if (isNotFoundError(e)) return base; // base name is free
      throw e;
    }
  }

  // ─── List view setup ──────────────────────────────────────────────────────

  private async _setupTaskListViews(listName: string): Promise<void> {
    // Default view: all task columns in logical order
    try {
      const dv = await this.sp.web.lists.getByTitle(listName).defaultView();
      const dvf = this.sp.web.lists.getByTitle(listName).views.getById(dv.Id).fields;
      await dvf.removeAll();
      for (const f of [
        'Title', 'Phase', 'Status', 'Priority', 'StartDate', 'DueDate',
        'AssignedToName', 'PercentComplete', 'IsMilestone',
        'Dependencies', 'Notes', 'SortOrder',
      ]) {
        try { await dvf.add(f); } catch { /* skip if field missing */ }
      }
    } catch (e) { console.warn('[SmartGantt] Default view setup (non-fatal):', e); }

    // Admin views for grouping and filtering
    const adminViews: Array<{ name: string; query: string; fields: string[] }> = [
      {
        name: 'By Phase',
        query: '<GroupBy Collapse="FALSE"><FieldRef Name="Phase"/></GroupBy>'
             + '<OrderBy><FieldRef Name="SortOrder"/></OrderBy>',
        fields: ['Title', 'Status', 'Priority', 'StartDate', 'DueDate', 'AssignedToName', 'PercentComplete'],
      },
      {
        name: 'By Status',
        query: '<GroupBy Collapse="FALSE"><FieldRef Name="Status"/></GroupBy>'
             + '<OrderBy><FieldRef Name="DueDate"/></OrderBy>',
        fields: ['Title', 'Phase', 'Priority', 'StartDate', 'DueDate', 'AssignedToName', 'PercentComplete'],
      },
      {
        name: 'By Assignee',
        query: '<GroupBy Collapse="FALSE"><FieldRef Name="AssignedToName"/></GroupBy>'
             + '<OrderBy><FieldRef Name="DueDate"/></OrderBy>',
        fields: ['Title', 'Phase', 'Status', 'Priority', 'StartDate', 'DueDate', 'PercentComplete'],
      },
      {
        name: 'Milestones',
        query: '<Where><Eq><FieldRef Name="IsMilestone"/><Value Type="Boolean">1</Value></Eq></Where>'
             + '<OrderBy><FieldRef Name="DueDate"/></OrderBy>',
        fields: ['Title', 'Phase', 'StartDate', 'DueDate', 'AssignedToName', 'Status'],
      },
    ];

    for (const vd of adminViews) {
      try {
        const created = await this.sp.web.lists.getByTitle(listName).views.add(vd.name, false, {
          ViewQuery: vd.query, RowLimit: 100,
        });
        const vf = this.sp.web.lists.getByTitle(listName).views.getById(created.data.Id).fields;
        await vf.removeAll();
        for (const f of vd.fields) {
          try { await vf.add(f); } catch { /* skip */ }
        }
      } catch (e) { console.warn(`[SmartGantt] View "${vd.name}" (non-fatal):`, e); }
    }
  }

  private async _setupMetaListView(): Promise<void> {
    try {
      const dv = await this.sp.web.lists.getByTitle(PROJECTS_LIST).defaultView();
      const dvf = this.sp.web.lists.getByTitle(PROJECTS_LIST).views.getById(dv.Id).fields;
      await dvf.removeAll();
      for (const f of [
        'Title', 'ProjectDescription', 'ProjectStatus',
        'ProjectStartDate', 'ProjectDueDate', 'ProjectManager', 'ProjectListName',
      ]) {
        try { await dvf.add(f); } catch { /* skip */ }
      }
    } catch (e) { console.warn('[SmartGantt] Meta-list view (non-fatal):', e); }
  }
}
