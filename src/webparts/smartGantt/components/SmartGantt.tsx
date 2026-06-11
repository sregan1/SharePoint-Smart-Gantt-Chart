import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  Spinner, SpinnerSize, Dialog, DialogType, DialogFooter,
  DefaultButton, PrimaryButton, MessageBar, MessageBarType,
} from '@fluentui/react';

import { SharePointService } from '../services/SharePointService';
import { exportTasksToExcel, renderGanttSVG, downloadPNG, exportToPowerPoint, exportPortfolioToExcel, exportPortfolioToPowerPoint } from '../services/ExportService';
import {
  IProject, ITask, IProjectTaskStats, ViewMode, ZoomLevel,
  IGanttDisplaySettings, DEFAULT_GANTT_SETTINGS, ITaskFilter, EMPTY_TASK_FILTER,
} from '../models';
import { filterTasks } from '../utils/filterUtils';
import { PortfolioView } from './views/PortfolioView';
import { Toolbar } from './toolbar/Toolbar';
import { GanttChart } from './gantt/GanttChart';
import { GanttSettings } from './gantt/GanttSettings';
import { ListView } from './views/ListView';
import { KanbanView } from './views/KanbanView';
import { ProjectPanel } from './panels/ProjectPanel';
import { TaskPanel } from './panels/TaskPanel';
import { ImportPanel } from './import/ImportPanel';
import { DashboardView } from './views/DashboardView';

import styles from './SmartGantt.module.scss';

export interface ISmartGanttProps {
  title: string;
  spService: SharePointService;
  context: WebPartContext;
}

// User preferences persisted per web part instance (localStorage).
interface IPersistedPrefs {
  viewMode?: ViewMode;
  zoomLevel?: ZoomLevel;
  ganttSettings?: Partial<IGanttDisplaySettings>;
  selectedProjectId?: number | null;
  showArchivedProjects?: boolean;
}

interface ISmartGanttState {
  projects: IProject[];
  selectedProject: IProject | null;
  tasks: ITask[];
  viewMode: ViewMode;
  zoomLevel: ZoomLevel;
  loading: boolean;
  tasksLoading: boolean;
  error: string | null;
  tasksError: string | null;
  saveError: string | null;
  showProjectPanel: boolean;
  editingProject: IProject | null;
  showTaskPanel: boolean;
  editingTask: ITask | null;
  deleteProjectConfirm: boolean;
  deleteTaskConfirm: ITask | null;
  archiveProjectConfirm: boolean;
  showArchivedProjects: boolean;
  scrollToToday: boolean;
  showImportPanel: boolean;
  showGanttSettings: boolean;
  ganttSettings: IGanttDisplaySettings;
  taskFilter: ITaskFilter;
  portfolioStats: Map<number, IProjectTaskStats> | null;
  portfolioLoading: boolean;
}

export default class SmartGantt extends React.Component<ISmartGanttProps, ISmartGanttState> {
  // Monotonic token so a slow task fetch for a previously selected project
  // can't overwrite the tasks of the currently selected one.
  private _taskLoadSeq = 0;
  private _prefsKey: string;

  constructor(props: ISmartGanttProps) {
    super(props);
    this._prefsKey = `SmartGantt_prefs_${props.context.instanceId}`;
    const prefs = this._loadPrefs();
    this.state = {
      projects: [],
      selectedProject: null,
      tasks: [],
      viewMode: prefs.viewMode || 'gantt',
      zoomLevel: prefs.zoomLevel || 'week',
      loading: true,
      tasksLoading: false,
      error: null,
      tasksError: null,
      saveError: null,
      showProjectPanel: false,
      editingProject: null,
      showTaskPanel: false,
      editingTask: null,
      deleteProjectConfirm: false,
      deleteTaskConfirm: null,
      archiveProjectConfirm: false,
      showArchivedProjects: prefs.showArchivedProjects || false,
      scrollToToday: false,
      showImportPanel: false,
      showGanttSettings: false,
      ganttSettings: { ...DEFAULT_GANTT_SETTINGS, ...prefs.ganttSettings },
      taskFilter: EMPTY_TASK_FILTER,
      portfolioStats: null,
      portfolioLoading: false,
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._loadProjects();
  }

  public componentDidUpdate(_prevProps: ISmartGanttProps, prevState: ISmartGanttState): void {
    const s = this.state;
    if (
      prevState.viewMode !== s.viewMode ||
      prevState.zoomLevel !== s.zoomLevel ||
      prevState.ganttSettings !== s.ganttSettings ||
      prevState.selectedProject?.id !== s.selectedProject?.id ||
      prevState.showArchivedProjects !== s.showArchivedProjects
    ) {
      this._savePrefs();
    }
  }

  // ─── Preference persistence ──────────────────────────────────────────────

  private _loadPrefs(): IPersistedPrefs {
    try {
      const raw = window.localStorage.getItem(this._prefsKey);
      return raw ? (JSON.parse(raw) as IPersistedPrefs) : {};
    } catch {
      return {};
    }
  }

  private _savePrefs(): void {
    try {
      const { viewMode, zoomLevel, ganttSettings, selectedProject, showArchivedProjects } = this.state;
      const prefs: IPersistedPrefs = {
        viewMode,
        zoomLevel,
        ganttSettings,
        selectedProjectId: selectedProject?.id ?? null,
        showArchivedProjects,
      };
      window.localStorage.setItem(this._prefsKey, JSON.stringify(prefs));
    } catch {
      // Storage unavailable (private mode, quota) — preferences just won't stick.
    }
  }

  // ─── Loading ─────────────────────────────────────────────────────────────

  private _errMessage(err: unknown, fallback: string): string {
    return err instanceof Error && err.message ? err.message : fallback;
  }

  private async _loadProjects(preferredId?: number): Promise<void> {
    try {
      this.setState({ loading: true, error: null });
      const projects = await this.props.spService.getProjects();
      const visible = projects.filter(p => !p.isArchived);
      // Keep the current (or persisted) selection when it still exists,
      // instead of always snapping back to the first project.
      const wantedId = preferredId
        ?? this.state.selectedProject?.id
        ?? this._loadPrefs().selectedProjectId
        ?? undefined;
      const selectedProject =
        projects.find(p => p.id === wantedId) || visible[0] || projects[0] || null;
      this.setState({ projects, selectedProject, loading: false });
      if (selectedProject) {
        await this._loadTasks(selectedProject.listName);
      } else {
        this.setState({ tasks: [] });
      }
    } catch (err) {
      this.setState({
        loading: false,
        error: this._errMessage(err, 'Failed to load projects. Check site permissions.'),
      });
    }
  }

  private async _loadTasks(listName: string): Promise<void> {
    const seq = ++this._taskLoadSeq;
    this.setState({ tasksLoading: true, tasksError: null });
    try {
      const tasks = await this.props.spService.getProjectTasks(listName);
      if (seq !== this._taskLoadSeq) return; // stale response — a newer load won
      this.setState({ tasks, tasksLoading: false });
    } catch (err) {
      if (seq !== this._taskLoadSeq) return;
      this.setState({
        tasks: [],
        tasksLoading: false,
        tasksError: this._errMessage(err, 'Failed to load tasks for this project.'),
      });
    }
  }

  private _handleSelectProject = async (project: IProject): Promise<void> => {
    this.setState({ selectedProject: project, tasks: [], taskFilter: EMPTY_TASK_FILTER });
    await this._loadTasks(project.listName);
  };

  private _handleViewChange = (viewMode: ViewMode): void => {
    this.setState({ viewMode });
    if (viewMode === 'portfolio' && this.state.portfolioStats === null) {
      void this._loadPortfolioStats();
    }
  };

  private _loadPortfolioStats = async (): Promise<void> => {
    this.setState({ portfolioLoading: true });
    const stats = await this.props.spService.getAllProjectStats(this.state.projects);
    this.setState({ portfolioStats: stats, portfolioLoading: false });
  };

  private _handleZoomChange = (zoomLevel: ZoomLevel): void => {
    this.setState({ zoomLevel });
  };

  private _handleScrollToToday = (): void => {
    this.setState({ scrollToToday: true }, () => {
      setTimeout(() => this.setState({ scrollToToday: false }), 100);
    });
  };

  // ─── Project panel ───────────────────────────────────────────────────────

  private _handleAddProject = (): void => {
    this.setState({ showProjectPanel: true, editingProject: null });
  };

  private _handleEditProject = (): void => {
    this.setState({ showProjectPanel: true, editingProject: this.state.selectedProject });
  };

  private _handleProjectSave = async (data: Partial<IProject>): Promise<void> => {
    const { editingProject } = this.state;
    try {
      if (editingProject) {
        await this.props.spService.updateProject(editingProject.id, data);
        await this._loadProjects(editingProject.id);
      } else {
        const created = await this.props.spService.createProject({
          title: data.title!,
          description: data.description || '',
          color: data.color || '#0078D4',
          startDate: data.startDate || '',
          dueDate: data.dueDate || '',
          status: data.status || 'Active',
        });
        this.setState(prev => ({
          projects: [...prev.projects, created],
          selectedProject: created,
          tasks: [],
        }));
      }
      this.setState({ showProjectPanel: false, editingProject: null, portfolioStats: null });
    } catch (err) {
      this.setState({
        saveError: `Could not save the project: ${this._errMessage(err, 'unknown error')}`,
      });
    }
  };

  private _handleDeleteProject = (): void => {
    this.setState({ deleteProjectConfirm: true });
  };

  private _confirmDeleteProject = async (): Promise<void> => {
    const { selectedProject } = this.state;
    if (!selectedProject) return;
    try {
      await this.props.spService.deleteProject(selectedProject.id, selectedProject.listName);
      const projects = this.state.projects.filter(p => p.id !== selectedProject.id);
      const visible = projects.filter(p => !p.isArchived);
      this.setState({
        projects,
        selectedProject: visible[0] || null,
        tasks: [],
        deleteProjectConfirm: false,
        portfolioStats: null,
      });
      if (visible[0]) {
        await this._loadTasks(visible[0].listName);
      }
    } catch (err) {
      this.setState({
        deleteProjectConfirm: false,
        saveError: `Could not delete the project: ${this._errMessage(err, 'unknown error')}`,
      });
    }
  };

  private _handleArchiveProject = (): void => {
    this.setState({ archiveProjectConfirm: true });
  };

  private _confirmArchiveProject = async (): Promise<void> => {
    const { selectedProject } = this.state;
    if (!selectedProject) return;
    try {
      await this.props.spService.archiveProject(selectedProject.id);
      const projects = this.state.projects.map(p =>
        p.id === selectedProject.id ? { ...p, isArchived: true } : p
      );
      const visible = projects.filter(p => !p.isArchived);
      this.setState({
        projects,
        selectedProject: visible[0] || null,
        tasks: [],
        archiveProjectConfirm: false,
        portfolioStats: null,
      });
      if (visible[0]) {
        await this._loadTasks(visible[0].listName);
      }
    } catch (err) {
      this.setState({
        archiveProjectConfirm: false,
        saveError: `Could not archive the project: ${this._errMessage(err, 'unknown error')}`,
      });
    }
  };

  private _handleUnarchiveProject = async (): Promise<void> => {
    const { selectedProject } = this.state;
    if (!selectedProject) return;
    try {
      await this.props.spService.unarchiveProject(selectedProject.id);
      const projects = this.state.projects.map(p =>
        p.id === selectedProject.id ? { ...p, isArchived: false } : p
      );
      const selected = projects.find(p => p.id === selectedProject.id) || null;
      this.setState({ projects, selectedProject: selected, portfolioStats: null });
    } catch (err) {
      this.setState({
        saveError: `Could not unarchive the project: ${this._errMessage(err, 'unknown error')}`,
      });
    }
  };

  // ─── Task panel ──────────────────────────────────────────────────────────

  private _handleAddTask = (): void => {
    this.setState({ showTaskPanel: true, editingTask: null });
  };

  private _handleEditTask = (task: ITask): void => {
    this.setState({ showTaskPanel: true, editingTask: task });
  };

  private _handleTaskSave = async (data: Partial<ITask>): Promise<void> => {
    const { selectedProject, editingTask } = this.state;
    if (!selectedProject) return;
    try {
      if (editingTask) {
        await this.props.spService.updateTask(selectedProject.listName, editingTask.id, data);
      } else {
        await this.props.spService.createTask(selectedProject.listName, {
          ...data,
          sortOrder: this.state.tasks.length,
        });
      }
      await this._loadTasks(selectedProject.listName);
      this.setState({ showTaskPanel: false, editingTask: null, portfolioStats: null });
    } catch (err) {
      this.setState({
        saveError: `Could not save the task: ${this._errMessage(err, 'unknown error')}`,
      });
    }
  };

  private _handleDeleteTask = (taskId: number): void => {
    const task = this.state.tasks.find(t => t.id === taskId) || null;
    if (task) this.setState({ deleteTaskConfirm: task });
  };

  private _confirmDeleteTask = async (): Promise<void> => {
    const { selectedProject, deleteTaskConfirm } = this.state;
    if (!selectedProject || !deleteTaskConfirm) return;
    this.setState({ deleteTaskConfirm: null });
    try {
      await this.props.spService.deleteTask(selectedProject.listName, deleteTaskConfirm.id);
      await this._loadTasks(selectedProject.listName);
      this.setState({ portfolioStats: null });
    } catch (err) {
      this.setState({
        saveError: `Could not delete the task: ${this._errMessage(err, 'unknown error')}`,
      });
    }
  };

  private _handleExportImage = async (): Promise<void> => {
    const { selectedProject, tasks, ganttSettings } = this.state;
    if (!selectedProject) return;
    const svg = renderGanttSVG(selectedProject, tasks, ganttSettings);
    await downloadPNG(svg, `${selectedProject.title} - Gantt Chart.png`, 2);
  };

  private _handleExportPowerPoint = async (): Promise<void> => {
    const { selectedProject, tasks, ganttSettings } = this.state;
    if (!selectedProject) return;
    try {
      await exportToPowerPoint(selectedProject, tasks, ganttSettings);
    } catch (err) {
      this.setState({
        saveError: `PowerPoint export failed: ${this._errMessage(err, 'unknown error')}`,
      });
    }
  };

  private _handlePortfolioExportExcel = (): void => {
    exportPortfolioToExcel(this.state.projects.filter(p => !p.isArchived), this.state.portfolioStats);
  };

  private _handlePortfolioExportPowerPoint = async (): Promise<void> => {
    try {
      await exportPortfolioToPowerPoint(this.state.projects.filter(p => !p.isArchived), this.state.portfolioStats);
    } catch (err) {
      this.setState({
        saveError: `Portfolio export failed: ${this._errMessage(err, 'unknown error')}`,
      });
    }
  };

  private _handleTaskUpdate = async (taskId: number, updates: Partial<ITask>): Promise<void> => {
    const { selectedProject } = this.state;
    if (!selectedProject) return;
    // Optimistic update
    this.setState(prev => ({
      tasks: prev.tasks.map(t => t.id === taskId ? { ...t, ...updates } : t),
    }));
    try {
      await this.props.spService.updateTask(selectedProject.listName, taskId, updates);
    } catch (err) {
      // Revert on error
      this.setState({
        saveError: `Could not update the task: ${this._errMessage(err, 'unknown error')}`,
      });
      await this._loadTasks(selectedProject.listName);
    }
  };

  public render(): React.ReactElement {
    const {
      projects, selectedProject, tasks, viewMode, zoomLevel,
      loading, tasksLoading, error, tasksError, saveError,
      showProjectPanel, editingProject,
      showTaskPanel, editingTask,
      deleteProjectConfirm, deleteTaskConfirm, archiveProjectConfirm, showArchivedProjects,
      scrollToToday, showImportPanel, showGanttSettings, ganttSettings,
      taskFilter, portfolioStats, portfolioLoading,
    } = this.state;

    const visibleProjects = showArchivedProjects ? projects : projects.filter(p => !p.isArchived);
    const hasArchivedProjects = projects.some(p => p.isArchived);

    // Derive autocomplete/filter lists from the full task set
    const knownPhases = Array.from(new Set(tasks.map(t => t.phase).filter(Boolean))).sort();
    const knownUsers = Array.from(new Set(tasks.map(t => t.assignedTo).filter(Boolean))).sort();

    const filteredTasks = filterTasks(tasks, taskFilter);

    return (
      <div className={styles.smartGantt}>
        <Toolbar
          projects={visibleProjects}
          selectedProject={selectedProject}
          viewMode={viewMode}
          zoomLevel={zoomLevel}
          onSelectProject={this._handleSelectProject}
          onViewChange={this._handleViewChange}
          onZoomChange={this._handleZoomChange}
          onScrollToToday={this._handleScrollToToday}
          onAddTask={this._handleAddTask}
          ganttSettings={ganttSettings}
          onAddProject={this._handleAddProject}
          onEditProject={this._handleEditProject}
          onDeleteProject={this._handleDeleteProject}
          onArchiveProject={this._handleArchiveProject}
          onUnarchiveProject={this._handleUnarchiveProject}
          showArchivedProjects={showArchivedProjects}
          hasArchivedProjects={hasArchivedProjects}
          onToggleShowArchived={() => this.setState(s => ({ showArchivedProjects: !s.showArchivedProjects }))}
          onImport={() => this.setState({ showImportPanel: true })}
          onExportExcel={() => { if (selectedProject) exportTasksToExcel(selectedProject, tasks); }}
          onExportImage={this._handleExportImage}
          onExportPowerPoint={this._handleExportPowerPoint}
          onPortfolioExportExcel={this._handlePortfolioExportExcel}
          onPortfolioExportPowerPoint={this._handlePortfolioExportPowerPoint}
          onOpenSettings={() => this.setState(s => ({ showGanttSettings: !s.showGanttSettings }))}
          showSettings={showGanttSettings}
          taskFilter={taskFilter}
          onFilterChange={f => this.setState({ taskFilter: f })}
          knownUsers={knownUsers}
          knownPhases={knownPhases}
          filteredCount={filteredTasks.length}
          totalCount={tasks.length}
        />

        {saveError && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={() => this.setState({ saveError: null })}
            dismissButtonAriaLabel="Dismiss"
          >
            {saveError}
          </MessageBar>
        )}

        <div className={styles.viewContainer}>
          {loading && (
            <div className={styles.loadingContainer}>
              <Spinner size={SpinnerSize.large} label="Loading projects…" />
            </div>
          )}

          {!loading && error && (
            <div className={styles.errorContainer}>
              <div className={styles.errorTitle}>⚠ Unable to load</div>
              <div className={styles.errorMessage}>{error}</div>
              <PrimaryButton text="Retry" onClick={() => void this._loadProjects()} />
            </div>
          )}

          {!loading && !error && !selectedProject && viewMode !== 'portfolio' && (
            <div className={styles.emptyState}>
              <div className={styles.emptyIcon}>📋</div>
              <div className={styles.emptyTitle}>No projects yet</div>
              <div className={styles.emptySubtitle}>
                Create your first project to start managing tasks with Gantt charts, lists, and Kanban boards.
              </div>
              <PrimaryButton text="+ Create First Project" onClick={this._handleAddProject} />
            </div>
          )}

          {!loading && !error && viewMode === 'portfolio' && (
            <PortfolioView
              projects={visibleProjects}
              statsMap={portfolioStats}
              loading={portfolioLoading}
              onSelectProject={(project) => {
                void this._handleSelectProject(project);
                this._handleViewChange('gantt');
              }}
              onAddProject={this._handleAddProject}
              onRefresh={this._loadPortfolioStats}
            />
          )}

          {!loading && !error && selectedProject && viewMode !== 'portfolio' && (
            <>
              {tasksLoading && (
                <div className={styles.loadingContainer} style={{ position: 'absolute', zIndex: 10, background: 'rgba(255,255,255,0.8)' }}>
                  <Spinner size={SpinnerSize.medium} label="Loading tasks…" />
                </div>
              )}

              {tasksError && (
                <MessageBar
                  messageBarType={MessageBarType.warning}
                  actions={
                    <DefaultButton
                      text="Retry"
                      onClick={() => { if (selectedProject) void this._loadTasks(selectedProject.listName); }}
                    />
                  }
                >
                  {tasksError}
                </MessageBar>
              )}

              {viewMode === 'gantt' && (
                <GanttChart
                  tasks={filteredTasks}
                  project={selectedProject}
                  zoomLevel={zoomLevel}
                  settings={ganttSettings}
                  scrollToToday={scrollToToday}
                  onEditTask={this._handleEditTask}
                  onDeleteTask={this._handleDeleteTask}
                  onTaskUpdate={this._handleTaskUpdate}
                  onAddTask={this._handleAddTask}
                  onImport={() => this.setState({ showImportPanel: true })}
                />
              )}

              {/* GanttSettings lives outside the viewMode block so the Options button
                  works from any view — the panel stays mounted and showGanttSettings
                  state is always reflected correctly. */}
              <GanttSettings
                isOpen={showGanttSettings}
                settings={ganttSettings}
                onChange={s => this.setState({ ganttSettings: s })}
                onDismiss={() => this.setState({ showGanttSettings: false })}
              />

              {viewMode === 'list' && (
                <ListView
                  tasks={filteredTasks}
                  project={selectedProject}
                  showHealthBadges={ganttSettings.showHealthBadges}
                  onEditTask={this._handleEditTask}
                  onDeleteTask={this._handleDeleteTask}
                  onTaskUpdate={this._handleTaskUpdate}
                  onAddTask={this._handleAddTask}
                />
              )}

              {viewMode === 'kanban' && (
                <KanbanView
                  tasks={filteredTasks}
                  project={selectedProject}
                  showHealthBadges={ganttSettings.showHealthBadges}
                  onEditTask={this._handleEditTask}
                  onDeleteTask={this._handleDeleteTask}
                  onTaskUpdate={this._handleTaskUpdate}
                  onAddTask={this._handleAddTask}
                />
              )}

              {viewMode === 'dashboard' && (
                <DashboardView
                  tasks={filteredTasks}
                  project={selectedProject}
                  onEditTask={this._handleEditTask}
                  onAddTask={this._handleAddTask}
                />
              )}
            </>
          )}
        </div>

        {/* Project Panel */}
        <ProjectPanel
          isOpen={showProjectPanel}
          project={editingProject}
          onSave={this._handleProjectSave}
          onDismiss={() => this.setState({ showProjectPanel: false, editingProject: null })}
        />

        {/* Task Panel */}
        {selectedProject && (
          <TaskPanel
            isOpen={showTaskPanel}
            task={editingTask}
            tasks={tasks}
            project={selectedProject}
            knownPhases={knownPhases}
            knownUsers={knownUsers}
            onSave={this._handleTaskSave}
            onDismiss={() => this.setState({ showTaskPanel: false, editingTask: null })}
          />
        )}

        {/* Import Panel */}
        {selectedProject && (
          <ImportPanel
            isOpen={showImportPanel}
            project={selectedProject}
            spService={this.props.spService}
            context={this.props.context}
            onDismiss={() => this.setState({ showImportPanel: false })}
            onImportComplete={() => {
              this.setState({ showImportPanel: false });
              if (selectedProject) void this._loadTasks(selectedProject.listName);
            }}
          />
        )}

        {/* Delete task confirm */}
        <Dialog
          hidden={!deleteTaskConfirm}
          onDismiss={() => this.setState({ deleteTaskConfirm: null })}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Delete Task?',
            subText: `"${deleteTaskConfirm?.title}" will be moved to the site recycle bin. You can restore it from there within 93 days.`,
          }}
        >
          <DialogFooter>
            <DefaultButton text="Cancel" onClick={() => this.setState({ deleteTaskConfirm: null })} />
            <PrimaryButton
              text="Delete"
              styles={{ root: { background: '#D13438', borderColor: '#D13438' } }}
              onClick={this._confirmDeleteTask}
            />
          </DialogFooter>
        </Dialog>

        {/* Archive project confirm */}
        <Dialog
          hidden={!archiveProjectConfirm}
          onDismiss={() => this.setState({ archiveProjectConfirm: false })}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Archive Project?',
            subText: `"${selectedProject?.title}" will be hidden from the project list and portfolio. You can restore it at any time via the project selector.`,
          }}
        >
          <DialogFooter>
            <DefaultButton text="Cancel" onClick={() => this.setState({ archiveProjectConfirm: false })} />
            <PrimaryButton text="Archive" onClick={this._confirmArchiveProject} />
          </DialogFooter>
        </Dialog>

        {/* Delete project confirm */}
        <Dialog
          hidden={!deleteProjectConfirm}
          onDismiss={() => this.setState({ deleteProjectConfirm: false })}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Send to Recycle Bin?',
            subText: `"${selectedProject?.title}" and all its tasks will be moved to the SharePoint recycle bin. You can restore them from the recycle bin within 93 days.`,
          }}
        >
          <DialogFooter>
            <DefaultButton text="Cancel" onClick={() => this.setState({ deleteProjectConfirm: false })} />
            <PrimaryButton
              text="Send to Recycle Bin"
              styles={{ root: { background: '#D13438', borderColor: '#D13438' } }}
              onClick={this._confirmDeleteProject}
            />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
}
