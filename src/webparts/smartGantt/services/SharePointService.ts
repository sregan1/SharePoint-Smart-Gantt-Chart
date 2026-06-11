import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/fields';
import '@pnp/sp/views';
import '@pnp/sp/batching';
import { IProject, ITask, IProjectTaskStats, ProjectStatus, TaskPriority, TaskStatus } from '../models';
import { computeTaskHealth, computeProjectHealth } from '../utils/healthUtils';
import { toDateOnly } from '../utils/dateUtils';

const PROJECTS_LIST = 'SmartGantt_Projects';

// Schedule dates are calendar days. We store them as UTC midnight so the same
// day is read back regardless of the viewer's timezone (PnPjs serializes Date
// objects for Edm.DateTime fields; raw ISO strings can fail in some tenants).
function toSPDate(s: string | undefined | null): Date | null {
  const dateOnly = toDateOnly(s);
  if (!dateOnly) return null;
  const [y, m, d] = dateOnly.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d));
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

export class SharePointService {
  private sp: SPFI;
  private projectsListEnsured = false;

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
    const list = this.sp.web.lists.getByTitle(PROJECTS_LIST);
    try {
      await list.fields.getByInternalNameOrTitle('IsArchived')();
    } catch {
      try {
        await list.fields.add('IsArchived', 8, {});
      } catch (e) {
        // Read-only users can't add fields; archiving simply won't be available.
        console.warn('[SmartGantt] IsArchived migration skipped:', e);
      }
    }
  }

  async getProjects(): Promise<IProject[]> {
    await this.ensureProjectsList();
    const items = await this.sp.web.lists
      .getByTitle(PROJECTS_LIST)
      .items.select(
        'Id', 'Title', 'ProjectListName', 'ProjectDescription', 'ProjectColor',
        'ProjectStartDate', 'ProjectDueDate', 'ProjectStatus', 'ProjectManager',
        'ProjectManagerEmail', 'Created', 'IsArchived'
      )
      .getAll();

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
    const listName = await this._buildListName(sanitized);

    await this._createProjectList(listName);

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
    await this.sp.web.lists.getByTitle(PROJECTS_LIST).items.getById(id).update(updates);
  }

  async archiveProject(id: number): Promise<void> {
    await this.sp.web.lists.getByTitle(PROJECTS_LIST).items.getById(id).update({ IsArchived: true });
  }

  async unarchiveProject(id: number): Promise<void> {
    await this.sp.web.lists.getByTitle(PROJECTS_LIST).items.getById(id).update({ IsArchived: false });
  }

  async deleteProject(id: number, listName: string): Promise<void> {
    await this.sp.web.lists.getByTitle(PROJECTS_LIST).items.getById(id).recycle();
    try {
      await this.sp.web.lists.getByTitle(listName).recycle();
    } catch {
      // list may not exist; ignore
    }
  }

  private async _createProjectList(listName: string): Promise<void> {
    await this.sp.web.lists.add(listName, 'Task list for Smart Gantt project', 100, false);

    // Brief pause so SharePoint fully provisions the list before we add fields.
    await new Promise<void>(resolve => setTimeout(resolve, 1500));

    // IsMilestone uses FieldTypeKind 8 (Boolean) via the generic add() because
    // addBoolean() in PnPjs 3.x passes the wrong SP type and triggers a 400.
    // All field adds go through a single REST batch — one round trip instead
    // of fifteen. Individual failures (duplicate field, etc.) are non-fatal.
    const [batchedSP, execute] = this.sp.batched();
    const list = batchedSP.web.lists.getByTitle(listName);
    const queue = (p: Promise<unknown>): void => {
      p.catch(e => console.warn('[SmartGantt] Field creation warning (non-fatal):', e));
    };

    queue(list.fields.addMultilineText('TaskDescription'));
    queue(list.fields.addDateTime('StartDate'));
    queue(list.fields.addDateTime('DueDate'));
    queue(list.fields.addChoice('Status', {
      Choices: ['Not Started', 'In Progress', 'Completed', 'On Hold', 'Cancelled'],
    }));
    queue(list.fields.addChoice('Priority', {
      Choices: ['Critical', 'High', 'Medium', 'Low'],
    }));
    queue(list.fields.addNumber('PercentComplete'));
    queue(list.fields.addNumber('ParentTaskId'));
    queue(list.fields.addText('Dependencies', { MaxLength: 500 }));
    queue(list.fields.addMultilineText('Notes'));
    queue(list.fields.addText('TaskColor', { MaxLength: 20 }));
    queue(list.fields.addNumber('SortOrder'));
    queue(list.fields.add('IsMilestone', 8));
    queue(list.fields.addText('Phase', { MaxLength: 100 }));
    queue(list.fields.addText('AssignedToName', { MaxLength: 255 }));
    queue(list.fields.addText('AssignedToEmail', { MaxLength: 255 }));

    await execute();

    await this._setupTaskListViews(listName);
  }

  // ─── Tasks ────────────────────────────────────────────────────────────────

  async getProjectTasks(listName: string): Promise<ITask[]> {
    const items = await this.sp.web.lists
      .getByTitle(listName)
      .items.select(
        'Id', 'Title', 'TaskDescription', 'StartDate', 'DueDate', 'Status', 'Priority',
        'AssignedToName', 'AssignedToEmail', 'PercentComplete', 'ParentTaskId', 'Dependencies',
        'Notes', 'TaskColor', 'SortOrder', 'IsMilestone', 'Phase', 'Created', 'Modified'
      )
      .getAll();

    return items
      .map(item => ({
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
        dependencies: item.Dependencies
          ? item.Dependencies.split(',').map(Number).filter(Boolean)
          : [],
        notes: item.Notes || '',
        color: item.TaskColor || '',
        sortOrder: item.SortOrder || 0,
        isMilestone: item.IsMilestone === true || item.IsMilestone === 1,
        phase: item.Phase || '',
        created: item.Created,
        modified: item.Modified,
      }))
      .sort((a, b) => a.sortOrder - b.sortOrder || a.id - b.id);
  }

  async getProjectTaskStats(project: IProject): Promise<IProjectTaskStats> {
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
      const items = await this.sp.web.lists
        .getByTitle(project.listName)
        .items.select('Status', 'Priority', 'PercentComplete', 'StartDate', 'DueDate', 'IsMilestone')
        .getAll();

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
    } catch {
      return empty;
    }
  }

  async getAllProjectStats(projects: IProject[]): Promise<Map<number, IProjectTaskStats>> {
    const results = await Promise.all(
      projects.map(p => this.getProjectTaskStats(p).then(stats => ({ id: p.id, stats })))
    );
    const map = new Map<number, IProjectTaskStats>();
    results.forEach(r => map.set(r.id, r.stats));
    return map;
  }

  async createTask(listName: string, task: Partial<ITask>): Promise<ITask> {
    const result = await this.sp.web.lists.getByTitle(listName).items.add({
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
      Dependencies: (task.dependencies || []).join(','),
      Notes: task.notes || '',
      TaskColor: task.color || '',
      SortOrder: task.sortOrder || 0,
      IsMilestone: task.isMilestone || false,
      Phase: task.phase || '',
    });

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
      notes: task.notes || '',
      color: task.color || '',
      sortOrder: task.sortOrder || 0,
      isMilestone: task.isMilestone || false,
      phase: task.phase || '',
      created: new Date().toISOString(),
      modified: new Date().toISOString(),
    };
  }

  async updateTask(listName: string, id: number, updates: Partial<ITask>): Promise<void> {
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
    if (updates.dependencies !== undefined) data.Dependencies = updates.dependencies.join(',');
    if (updates.notes !== undefined) data.Notes = updates.notes;
    if (updates.color !== undefined) data.TaskColor = updates.color;
    if (updates.sortOrder !== undefined) data.SortOrder = updates.sortOrder;
    if (updates.isMilestone !== undefined) data.IsMilestone = updates.isMilestone;
    if (updates.phase !== undefined) data.Phase = updates.phase;
    await this.sp.web.lists.getByTitle(listName).items.getById(id).update(data);
  }

  async deleteTask(listName: string, id: number): Promise<void> {
    const list = this.sp.web.lists.getByTitle(listName);

    // Promote sub-tasks to top level first, so they don't become invisible
    // orphans (views only render sub-tasks under an existing parent).
    try {
      const children = await list.items.select('Id').filter(`ParentTaskId eq ${id}`).top(500)();
      if (children.length > 0) {
        const [batchedSP, execute] = this.sp.batched();
        const batchedList = batchedSP.web.lists.getByTitle(listName);
        children.forEach(c => {
          batchedList.items.getById(c.Id).update({ ParentTaskId: 0 })
            .catch(e => console.warn('[SmartGantt] Sub-task promotion warning:', e));
        });
        await execute();
      }
    } catch (e) {
      console.warn('[SmartGantt] Orphan cleanup (non-fatal):', e);
    }

    // Recycle (not delete) so the task can be restored from the recycle bin.
    await list.items.getById(id).recycle();
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
        } catch {
          return candidate; // free
        }
      }
      return `${base}_${Date.now().toString(36).toUpperCase()}`; // extremely unlikely fallback
    } catch {
      return base; // base name is free
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
