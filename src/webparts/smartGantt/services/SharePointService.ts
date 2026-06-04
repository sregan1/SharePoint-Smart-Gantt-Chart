import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/views';
import { IProject, ITask, ProjectStatus, TaskPriority, TaskStatus } from '../models';

const PROJECTS_LIST = 'SmartGantt_Projects';

// PnPjs serializes Date objects correctly for SharePoint Edm.DateTime fields.
// Passing raw ISO strings can fail in some tenant configurations.
function toSPDate(s: string | undefined | null): Date | null {
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

export class SharePointService {
  private sp: SPFI;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ─── Projects ────────────────────────────────────────────────────────────

  async ensureProjectsList(): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(PROJECTS_LIST)();
    } catch {
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
      await this._setupMetaListView();
    }
  }

  async getProjects(): Promise<IProject[]> {
    await this.ensureProjectsList();
    const items = await this.sp.web.lists
      .getByTitle(PROJECTS_LIST)
      .items.select(
        'Id', 'Title', 'ProjectListName', 'ProjectDescription', 'ProjectColor',
        'ProjectStartDate', 'ProjectDueDate', 'ProjectStatus', 'ProjectManager',
        'ProjectManagerEmail', 'Created'
      )
      .orderBy('Created', true)
      .top(500)();

    return items.map(item => ({
      id: item.Id,
      title: item.Title,
      listName: item.ProjectListName || '',
      description: item.ProjectDescription || '',
      color: item.ProjectColor || '#0078D4',
      startDate: item.ProjectStartDate || '',
      dueDate: item.ProjectDueDate || '',
      status: (item.ProjectStatus || 'Active') as ProjectStatus,
      projectManager: item.ProjectManager || '',
      projectManagerEmail: item.ProjectManagerEmail || '',
      created: item.Created,
    }));
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

    const sanitized = data.title.replace(/[^a-zA-Z0-9]/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '');
    const listName = await this._buildListName(sanitized);

    await this._createProjectList(listName);

    const result = await this.sp.web.lists.getByTitle(PROJECTS_LIST).items.add({
      Title: data.title,
      ProjectListName: listName,
      ProjectDescription: data.description,
      ProjectColor: data.color,
      ProjectStartDate: toSPDate(data.startDate),
      ProjectDueDate: toSPDate(data.dueDate),
      ProjectStatus: data.status,
    });

    return {
      id: result.data.Id,
      title: data.title,
      listName,
      description: data.description,
      color: data.color,
      startDate: data.startDate,
      dueDate: data.dueDate,
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
    await this.sp.web.lists.getByTitle(PROJECTS_LIST).items.getById(id).update(updates);
  }

  async deleteProject(id: number, listName: string): Promise<void> {
    await this.sp.web.lists.getByTitle(PROJECTS_LIST).items.getById(id).delete();
    try {
      await this.sp.web.lists.getByTitle(listName).delete();
    } catch {
      // list may not exist; ignore
    }
  }

  private async _createProjectList(listName: string): Promise<void> {
    await this.sp.web.lists.add(listName, 'Task list for Smart Gantt project', 100, false);

    // Brief pause so SharePoint fully provisions the list before we add fields.
    await new Promise<void>(resolve => setTimeout(resolve, 1500));

    const list = this.sp.web.lists.getByTitle(listName);

    // Field definitions: [title, addMethod, properties?]
    // IsMilestone uses FieldTypeKind 8 (Boolean) via the generic add() because
    // addBoolean() in PnPjs 3.x passes the wrong SP type and triggers a 400.
    const fields: Array<() => Promise<unknown>> = [
      () => list.fields.addMultilineText('TaskDescription'),
      () => list.fields.addDateTime('StartDate'),
      () => list.fields.addDateTime('DueDate'),
      () => list.fields.addChoice('Status', {
        Choices: ['Not Started', 'In Progress', 'Completed', 'On Hold', 'Cancelled'],
      }),
      () => list.fields.addChoice('Priority', {
        Choices: ['Critical', 'High', 'Medium', 'Low'],
      }),
      () => list.fields.addNumber('PercentComplete'),
      () => list.fields.addNumber('ParentTaskId'),
      () => list.fields.addText('Dependencies', { MaxLength: 500 }),
      () => list.fields.addMultilineText('Notes'),
      () => list.fields.addText('TaskColor', { MaxLength: 20 }),
      () => list.fields.addNumber('SortOrder'),
      // FieldTypeKind 8 = Boolean. Using add() directly because addBoolean()
      // in PnPjs 3.x passes the wrong __metadata type and triggers a SharePoint 400.
      () => list.fields.add('IsMilestone', 8),
      () => list.fields.addText('Phase', { MaxLength: 100 }),
      () => list.fields.addText('AssignedToName', { MaxLength: 255 }),
      () => list.fields.addText('AssignedToEmail', { MaxLength: 255 }),
    ];

    for (const addField of fields) {
      try {
        await addField();
      } catch (e) {
        // Log but continue — a duplicate or pre-existing field is non-fatal.
        console.warn('[SmartGantt] Field creation warning (non-fatal):', e);
      }
    }

    await this._setupTaskListViews(listName);
  }

  // ─── Tasks ────────────────────────────────────────────────────────────────

  async getProjectTasks(listName: string): Promise<ITask[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(listName)
        .items.select(
          'Id', 'Title', 'TaskDescription', 'StartDate', 'DueDate', 'Status', 'Priority',
          'AssignedToName', 'AssignedToEmail', 'PercentComplete', 'ParentTaskId', 'Dependencies',
          'Notes', 'TaskColor', 'SortOrder', 'IsMilestone', 'Phase', 'Created', 'Modified'
        )
        .orderBy('SortOrder', true)
        .top(1000)();

      return items.map(item => ({
        id: item.Id,
        title: item.Title,
        description: item.TaskDescription || '',
        startDate: item.StartDate || '',
        dueDate: item.DueDate || '',
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
      }));
    } catch {
      return [];
    }
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
      startDate: task.startDate || '',
      dueDate: task.dueDate || '',
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
    await this.sp.web.lists.getByTitle(listName).items.getById(id).delete();
  }

  async reorderTasks(listName: string, taskIds: number[]): Promise<void> {
    await Promise.all(
      taskIds.map((id, index) =>
        this.sp.web.lists.getByTitle(listName).items.getById(id).update({ SortOrder: index })
      )
    );
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
