import * as React from 'react';
import * as strings from 'SmartGanttWebPartStrings';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  Spinner, SpinnerSize, Stack, Dialog, DialogType, DialogFooter,
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
import { formatString } from './localeUtils';

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
  showImportAsProject: boolean;
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
  // Same idea for portfolio stats — a fetch kicked off before the project
  // list finished loading (over an empty array) must not clobber a later,
  // correct fetch that raced ahead of it.
  private _portfolioLoadSeq = 0;
  private _prefsKey: string;

  // Manual memoization for a class component: these derived values are
  // recomputed on every render otherwise (a class component has no useMemo),
  // including renders triggered by unrelated state like a callout toggling —
  // and filteredTasks in particular feeds the Gantt/List/Kanban view, so
  // recomputing it needlessly means those re-render with new array
  // identities too, defeating any memoization further down.
  private _visibleProjectsCache: { projects: IProject[]; showArchived: boolean; result: IProject[] } | null = null;
  private _knownListsCache: { tasks: ITask[]; phases: string[]; users: string[] } | null = null;
  private _filteredTasksCache: { tasks: ITask[]; filter: ITaskFilter; result: ITask[] } | null = null;

  private _getVisibleProjects(projects: IProject[], showArchived: boolean): IProject[] {
    const c = this._visibleProjectsCache;
    if (c && c.projects === projects && c.showArchived === showArchived) return c.result;
    const result = showArchived ? projects : projects.filter(p => !p.isArchived);
    this._visibleProjectsCache = { projects, showArchived, result };
    return result;
  }

  private _getKnownLists(tasks: ITask[]): { phases: string[]; users: string[] } {
    const c = this._knownListsCache;
    if (c && c.tasks === tasks) return c;
    const phases = Array.from(new Set(tasks.map(t => t.phase).filter(Boolean))).sort();
    const users = Array.from(new Set(tasks.map(t => t.assignedTo).filter(Boolean))).sort();
    const result = { tasks, phases, users };
    this._knownListsCache = result;
    return result;
  }

  private _getFilteredTasks(tasks: ITask[], filter: ITaskFilter): ITask[] {
    const c = this._filteredTasksCache;
    if (c && c.tasks === tasks && c.filter === filter) return c.result;
    const result = filterTasks(tasks, filter);
    this._filteredTasksCache = { tasks, filter, result };
    return result;
  }

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
      showImportAsProject: false,
      showGanttSettings: false,
      ganttSettings: { ...DEFAULT_GANTT_SETTINGS, ...prefs.ganttSettings },
      taskFilter: EMPTY_TASK_FILTER,
      portfolioStats: null,
      portfolioLoading: false,
    };
  }

  private _scrollToTodayTimer: ReturnType<typeof setTimeout> | null = null;
  private _isMounted = false;

  public async componentDidMount(): Promise<void> {
    this._isMounted = true;
    await this._loadProjects();
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
    if (this._scrollToTodayTimer !== null) clearTimeout(this._scrollToTodayTimer);
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
      if (!this._isMounted) return;
      const visible = projects.filter(p => !p.isArchived);
      // Keep the current (or persisted) selection when it still exists,
      // instead of always snapping back to the first project.
      const wantedId = preferredId
        ?? this.state.selectedProject?.id
        ?? this._loadPrefs().selectedProjectId
        ?? undefined;
      // Prefer a visible (non-archived) match so a persisted/previous
      // selection that's since been archived doesn't restore a project the
      // toolbar's project list isn't currently showing.
      const selectedProject =
        visible.find(p => p.id === wantedId)
        || (this.state.showArchivedProjects ? projects.find(p => p.id === wantedId) : undefined)
        || visible[0]
        || null;
      this.setState({ projects, selectedProject, loading: false });
      if (selectedProject) {
        await this._loadTasks(selectedProject.listName);
      } else {
        this.setState({ tasks: [] });
      }
      // Persisted view can be 'portfolio' on first mount, or the project
      // list can change (create/delete/archive) while portfolio is active —
      // either way, refresh stats against the just-loaded project list. This
      // also resolves any stats fetch that raced ahead using an empty/stale
      // project array (see the seq guard in _loadPortfolioStats).
      if (this._isMounted && this.state.viewMode === 'portfolio') {
        void this._loadPortfolioStats();
      }
    } catch (err) {
      if (!this._isMounted) return;
      this.setState({
        loading: false,
        error: this._errMessage(err, strings.SmartGantt_FailedLoadProjects),
      });
    }
  }

  private async _loadTasks(listName: string): Promise<void> {
    const seq = ++this._taskLoadSeq;
    this.setState({ tasksLoading: true, tasksError: null });
    try {
      const tasks = await this.props.spService.getProjectTasks(listName);
      if (seq !== this._taskLoadSeq || !this._isMounted) return; // stale response — a newer load won
      this.setState({ tasks, tasksLoading: false });
    } catch (err) {
      if (seq !== this._taskLoadSeq || !this._isMounted) return;
      this.setState({
        tasks: [],
        tasksLoading: false,
        tasksError: this._errMessage(err, strings.SmartGantt_FailedLoadTasks),
      });
    }
  }

  private _handleSelectProject = async (project: IProject): Promise<void> => {
    this.setState(prevState => ({
      selectedProject: project,
      tasks: [],
      taskFilter: EMPTY_TASK_FILTER,
      viewMode: prevState.viewMode === 'portfolio' ? 'gantt' : prevState.viewMode,
    }));
    await this._loadTasks(project.listName);
  };

  private _handleViewChange = (viewMode: ViewMode): void => {
    this.setState({ viewMode });
    if (viewMode === 'portfolio' && this.state.portfolioStats === null) {
      void this._loadPortfolioStats();
    }
  };

  private _loadPortfolioStats = async (): Promise<void> => {
    const seq = ++this._portfolioLoadSeq;
    this.setState({ portfolioLoading: true });
    try {
      const stats = await this.props.spService.getAllProjectStats(this.state.projects);
      if (seq !== this._portfolioLoadSeq || !this._isMounted) return;
      this.setState({ portfolioStats: stats, portfolioLoading: false });
    } catch (err) {
      if (seq !== this._portfolioLoadSeq || !this._isMounted) return;
      this.setState({
        portfolioLoading: false,
        saveError: this._errMessage(err, strings.SmartGantt_FailedLoadPortfolioStats),
      });
    }
  };

  private _handleZoomChange = (zoomLevel: ZoomLevel): void => {
    this.setState({ zoomLevel });
  };

  private _handleScrollToToday = (): void => {
    this.setState({ scrollToToday: true }, () => {
      if (this._scrollToTodayTimer !== null) clearTimeout(this._scrollToTodayTimer);
      this._scrollToTodayTimer = setTimeout(() => {
        this._scrollToTodayTimer = null;
        if (this._isMounted) this.setState({ scrollToToday: false });
      }, 100);
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
          // A filter carried over from the previous project could hide every
          // task in this brand-new one, with no visible way to tell why.
          taskFilter: EMPTY_TASK_FILTER,
        }));
      }
      this.setState({ showProjectPanel: false, editingProject: null, portfolioStats: null });
    } catch (err) {
      const message = formatString(strings.SmartGantt_CouldNotSaveProject, { message: this._errMessage(err, 'unknown error') });
      this.setState({ saveError: message });
      // Re-thrown so the still-open panel can show the error inline instead
      // of it only appearing behind the modal overlay.
      throw new Error(message);
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
        saveError: formatString(strings.SmartGantt_CouldNotDeleteProject, { message: this._errMessage(err, 'unknown error') }),
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
        saveError: formatString(strings.SmartGantt_CouldNotArchiveProject, { message: this._errMessage(err, 'unknown error') }),
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
        saveError: formatString(strings.SmartGantt_CouldNotUnarchiveProject, { message: this._errMessage(err, 'unknown error') }),
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
      const message = formatString(strings.SmartGantt_CouldNotSaveTask, { message: this._errMessage(err, 'unknown error') });
      this.setState({ saveError: message });
      // Re-thrown so the still-open panel can show the error inline instead
      // of it only appearing behind the modal overlay.
      throw new Error(message);
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
        saveError: formatString(strings.SmartGantt_CouldNotDeleteTask, { message: this._errMessage(err, 'unknown error') }),
      });
    }
  };

  private _handleExportImage = async (): Promise<void> => {
    const { selectedProject, tasks, ganttSettings } = this.state;
    if (!selectedProject) return;
    try {
      const svg = renderGanttSVG(selectedProject, tasks, ganttSettings);
      await downloadPNG(svg, `${selectedProject.title} - Gantt Chart.png`, 2);
    } catch (err) {
      this.setState({
        saveError: formatString(strings.SmartGantt_ImageExportFailed, { message: this._errMessage(err, 'unknown error') }),
      });
    }
  };

  // Hoisted from inline render-body arrows so Toolbar/FilterBar (wrapped in
  // React.memo) get stable prop identities and can actually skip re-rendering.
  private _handleToggleShowArchived = (): void => {
    this.setState(s => ({ showArchivedProjects: !s.showArchivedProjects }));
  };

  private _handleShowImportPanel = (): void => {
    this.setState({ showImportPanel: true });
  };

  private _handleShowImportAsProject = (): void => {
    this.setState({ showImportAsProject: true });
  };

  private _handleExportExcel = (): void => {
    const { selectedProject, tasks } = this.state;
    if (selectedProject) exportTasksToExcel(selectedProject, tasks);
  };

  private _handleToggleGanttSettings = (): void => {
    this.setState(s => ({ showGanttSettings: !s.showGanttSettings }));
  };

  private _handleFilterChange = (f: ITaskFilter): void => {
    this.setState({ taskFilter: f });
  };

  private _handleExportPowerPoint = async (): Promise<void> => {
    const { selectedProject, tasks, ganttSettings } = this.state;
    if (!selectedProject) return;
    try {
      await exportToPowerPoint(selectedProject, tasks, ganttSettings);
    } catch (err) {
      this.setState({
        saveError: formatString(strings.SmartGantt_PowerPointExportFailed, { message: this._errMessage(err, 'unknown error') }),
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
        saveError: formatString(strings.SmartGantt_PortfolioExportFailed, { message: this._errMessage(err, 'unknown error') }),
      });
    }
  };

  private _handleTaskUpdate = async (taskId: number, updates: Partial<ITask>): Promise<void> => {
    const { selectedProject } = this.state;
    if (!selectedProject) return;
    // Task IDs are only unique within a project's list — capture which
    // project this update belongs to so a project switch mid-flight can't
    // patch/reload the wrong project's tasks.
    const projectId = selectedProject.id;
    // Optimistic update
    this.setState(prev => {
      if (prev.selectedProject?.id !== projectId) return null;
      return { tasks: prev.tasks.map(t => t.id === taskId ? { ...t, ...updates } : t) };
    });
    try {
      await this.props.spService.updateTask(selectedProject.listName, taskId, updates);
    } catch (err) {
      if (this.state.selectedProject?.id !== projectId) return;
      // Revert on error
      this.setState({
        saveError: formatString(strings.SmartGantt_CouldNotUpdateTask, { message: this._errMessage(err, 'unknown error') }),
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
      scrollToToday, showImportPanel, showImportAsProject, showGanttSettings, ganttSettings,
      taskFilter, portfolioStats, portfolioLoading,
    } = this.state;

    const visibleProjects = this._getVisibleProjects(projects, showArchivedProjects);
    const hasArchivedProjects = projects.some(p => p.isArchived);

    // Derive autocomplete/filter lists from the full task set
    const { phases: knownPhases, users: knownUsers } = this._getKnownLists(tasks);

    const filteredTasks = this._getFilteredTasks(tasks, taskFilter);

    return (
      <div className={styles.smartGantt}>
        {this.props.title && <div className={styles.webPartTitle}>{this.props.title}</div>}
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
          onToggleShowArchived={this._handleToggleShowArchived}
          onImport={this._handleShowImportPanel}
          onImportAsProject={this._handleShowImportAsProject}
          onExportExcel={this._handleExportExcel}
          onExportImage={this._handleExportImage}
          onExportPowerPoint={this._handleExportPowerPoint}
          onPortfolioExportExcel={this._handlePortfolioExportExcel}
          onPortfolioExportPowerPoint={this._handlePortfolioExportPowerPoint}
          onOpenSettings={this._handleToggleGanttSettings}
          showSettings={showGanttSettings}
          taskFilter={taskFilter}
          onFilterChange={this._handleFilterChange}
          knownUsers={knownUsers}
          knownPhases={knownPhases}
          filteredCount={filteredTasks.length}
          totalCount={tasks.length}
        />

        {saveError && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={() => this.setState({ saveError: null })}
            dismissButtonAriaLabel={strings.SmartGantt_DismissAriaLabel}
          >
            {saveError}
          </MessageBar>
        )}

        <div className={styles.viewContainer}>
          {loading && (
            <div className={styles.loadingContainer}>
              <Spinner size={SpinnerSize.large} label={strings.SmartGantt_LoadingProjects} />
            </div>
          )}

          {!loading && error && (
            <div className={styles.errorContainer}>
              <div className={styles.errorTitle}>⚠ {strings.SmartGantt_UnableToLoad}</div>
              <div className={styles.errorMessage}>{error}</div>
              <PrimaryButton text={strings.SmartGantt_RetryButton} onClick={() => void this._loadProjects()} />
            </div>
          )}

          {!loading && !error && !selectedProject && viewMode !== 'portfolio' && (
            <div className={styles.emptyState}>
              <div className={styles.emptyIcon}>📋</div>
              <div className={styles.emptyTitle}>{strings.SmartGantt_NoProjectsYet}</div>
              <div className={styles.emptySubtitle}>
                {strings.SmartGantt_NoProjectsSubtitle}
              </div>
              <Stack horizontal tokens={{ childrenGap: 10 }} horizontalAlign="center">
                <PrimaryButton text={strings.SmartGantt_CreateFirstProjectButton} onClick={this._handleAddProject} />
                <DefaultButton
                  text={strings.SmartGantt_ImportFromExcelButton}
                  onClick={() => this.setState({ showImportAsProject: true })}
                />
              </Stack>
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
                  <Spinner size={SpinnerSize.medium} label={strings.SmartGantt_LoadingTasks} />
                </div>
              )}

              {tasksError && (
                <MessageBar
                  messageBarType={MessageBarType.warning}
                  actions={
                    <DefaultButton
                      text={strings.SmartGantt_RetryButton}
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

        {/* Import Panel — import into existing project */}
        {selectedProject && (
          <ImportPanel
            isOpen={showImportPanel}
            project={selectedProject}
            existingTasks={tasks}
            spService={this.props.spService}
            context={this.props.context}
            onDismiss={() => this.setState({ showImportPanel: false })}
            onImportComplete={() => {
              this.setState({ showImportPanel: false });
              if (selectedProject) void this._loadTasks(selectedProject.listName);
            }}
          />
        )}

        {/* Import Panel — create a new project from file */}
        <ImportPanel
          isOpen={showImportAsProject}
          spService={this.props.spService}
          context={this.props.context}
          onDismiss={() => this.setState({ showImportAsProject: false })}
          onImportComplete={(newProject) => {
            this.setState({ showImportAsProject: false });
            if (newProject) void this._loadProjects(newProject.id);
          }}
        />

        {/* Delete task confirm */}
        <Dialog
          hidden={!deleteTaskConfirm}
          onDismiss={() => this.setState({ deleteTaskConfirm: null })}
          dialogContentProps={{
            type: DialogType.normal,
            title: strings.SmartGantt_DeleteTaskDialogTitle,
            subText: formatString(strings.SmartGantt_DeleteTaskDialogSubtext, { taskTitle: deleteTaskConfirm?.title || '' }),
          }}
        >
          <DialogFooter>
            <DefaultButton text={strings.SmartGantt_CancelButton} onClick={() => this.setState({ deleteTaskConfirm: null })} />
            <PrimaryButton
              text={strings.SmartGantt_DeleteButton}
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
            title: strings.SmartGantt_ArchiveProjectDialogTitle,
            subText: formatString(strings.SmartGantt_ArchiveProjectDialogSubtext, { projectTitle: selectedProject?.title || '' }),
          }}
        >
          <DialogFooter>
            <DefaultButton text={strings.SmartGantt_CancelButton} onClick={() => this.setState({ archiveProjectConfirm: false })} />
            <PrimaryButton text={strings.SmartGantt_ArchiveButton} onClick={this._confirmArchiveProject} />
          </DialogFooter>
        </Dialog>

        {/* Delete project confirm */}
        <Dialog
          hidden={!deleteProjectConfirm}
          onDismiss={() => this.setState({ deleteProjectConfirm: false })}
          dialogContentProps={{
            type: DialogType.normal,
            title: strings.SmartGantt_SendToRecycleBinDialogTitle,
            subText: formatString(strings.SmartGantt_SendToRecycleBinDialogSubtext, { projectTitle: selectedProject?.title || '' }),
          }}
        >
          <DialogFooter>
            <DefaultButton text={strings.SmartGantt_CancelButton} onClick={() => this.setState({ deleteProjectConfirm: false })} />
            <PrimaryButton
              text={strings.SmartGantt_SendToRecycleBinButton}
              styles={{ root: { background: '#D13438', borderColor: '#D13438' } }}
              onClick={this._confirmDeleteProject}
            />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
}
