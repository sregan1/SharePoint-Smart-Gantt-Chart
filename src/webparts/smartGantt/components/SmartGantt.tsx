import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Spinner, SpinnerSize, Dialog, DialogType, DialogFooter, DefaultButton, PrimaryButton } from '@fluentui/react';

import { SharePointService } from '../services/SharePointService';
import { exportTasksToExcel, renderGanttSVG, downloadPNG, exportToPowerPoint, exportPortfolioToExcel, exportPortfolioToPowerPoint } from '../services/ExportService';
import { IProject, ITask, IProjectTaskStats, ViewMode, ZoomLevel, IGanttDisplaySettings, DEFAULT_GANTT_SETTINGS } from '../models';
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

interface ISmartGanttState {
  projects: IProject[];
  selectedProject: IProject | null;
  tasks: ITask[];
  viewMode: ViewMode;
  zoomLevel: ZoomLevel;
  loading: boolean;
  tasksLoading: boolean;
  error: string | null;
  showProjectPanel: boolean;
  editingProject: IProject | null;
  showTaskPanel: boolean;
  editingTask: ITask | null;
  deleteProjectConfirm: boolean;
  archiveProjectConfirm: boolean;
  showArchivedProjects: boolean;
  scrollToToday: boolean;
  showImportPanel: boolean;
  showGanttSettings: boolean;
  ganttSettings: IGanttDisplaySettings;
  portfolioStats: Map<number, IProjectTaskStats> | null;
  portfolioLoading: boolean;
}

export default class SmartGantt extends React.Component<ISmartGanttProps, ISmartGanttState> {
  constructor(props: ISmartGanttProps) {
    super(props);
    this.state = {
      projects: [],
      selectedProject: null,
      tasks: [],
      viewMode: 'gantt',
      zoomLevel: 'week',
      loading: true,
      tasksLoading: false,
      error: null,
      showProjectPanel: false,
      editingProject: null,
      showTaskPanel: false,
      editingTask: null,
      deleteProjectConfirm: false,
      archiveProjectConfirm: false,
      showArchivedProjects: false,
      scrollToToday: false,
      showImportPanel: false,
      showGanttSettings: false,
      ganttSettings: DEFAULT_GANTT_SETTINGS,
      portfolioStats: null,
      portfolioLoading: false,
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._loadProjects();
  }

  private async _loadProjects(): Promise<void> {
    try {
      this.setState({ loading: true, error: null });
      const projects = await this.props.spService.getProjects();
      const selectedProject = projects.length > 0 ? projects[0] : null;
      this.setState({ projects, selectedProject, loading: false });
      if (selectedProject) {
        await this._loadTasks(selectedProject.listName);
      }
    } catch (_err) {
      this.setState({
        loading: false,
        error: _err instanceof Error ? _err.message : 'Failed to load projects. Check site permissions.',
      });
    }
  }

  private async _loadTasks(listName: string): Promise<void> {
    try {
      this.setState({ tasksLoading: true });
      const tasks = await this.props.spService.getProjectTasks(listName);
      this.setState({ tasks, tasksLoading: false });
    } catch (_err) {
      this.setState({ tasksLoading: false });
    }
  }

  private _handleSelectProject = async (project: IProject): Promise<void> => {
    this.setState({ selectedProject: project, tasks: [] });
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
        this.setState({ showProjectPanel: false, editingProject: null });
        return;
      }
      await this._loadProjects();
      this.setState({ showProjectPanel: false, editingProject: null, portfolioStats: null });
    } catch (_err) {
      console.error('Failed to save project', _err);
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
    } catch (_err) {
      this.setState({ deleteProjectConfirm: false });
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
    } catch (_err) {
      this.setState({ archiveProjectConfirm: false });
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
      this.setState({ projects, portfolioStats: null });
    } catch (_err) {
      console.error('Failed to unarchive project', _err);
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
    } catch (_err) {
      console.error('Failed to save task', _err);
    }
  };

  private _handleDeleteTask = async (taskId: number): Promise<void> => {
    const { selectedProject } = this.state;
    if (!selectedProject) return;
    try {
      await this.props.spService.deleteTask(selectedProject.listName, taskId);
      await this._loadTasks(selectedProject.listName);
      this.setState({ portfolioStats: null });
    } catch (_err) {
      console.error('Failed to delete task', _err);
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
      console.error('PowerPoint export failed', err);
      alert('Export failed. Please try again.');
    }
  };

  private _handlePortfolioExportExcel = (): void => {
    exportPortfolioToExcel(this.state.projects, this.state.portfolioStats);
  };

  private _handlePortfolioExportPowerPoint = async (): Promise<void> => {
    try {
      await exportPortfolioToPowerPoint(this.state.projects, this.state.portfolioStats);
    } catch (err) {
      console.error('Portfolio PowerPoint export failed', err);
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
    } catch (_err) {
      // Revert on error
      await this._loadTasks(selectedProject.listName);
    }
  };

  public render(): React.ReactElement {
    const {
      projects, selectedProject, tasks, viewMode, zoomLevel,
      loading, tasksLoading, error,
      showProjectPanel, editingProject,
      showTaskPanel, editingTask,
      deleteProjectConfirm, archiveProjectConfirm, showArchivedProjects,
      scrollToToday, showImportPanel, showGanttSettings, ganttSettings,
      portfolioStats, portfolioLoading,
    } = this.state;

    const visibleProjects = showArchivedProjects ? projects : projects.filter(p => !p.isArchived);
    const hasArchivedProjects = projects.some(p => p.isArchived);

    // Derive autocomplete lists from current tasks
    const knownPhases = Array.from(new Set(tasks.map(t => t.phase).filter(Boolean))).sort() as string[];
    const knownUsers = Array.from(new Set(tasks.map(t => t.assignedTo).filter(Boolean))).sort() as string[];

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
        />

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
              <PrimaryButton text="Retry" onClick={this._loadProjects.bind(this)} />
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
                this._handleSelectProject(project);
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

              {viewMode === 'gantt' && (
                <GanttChart
                  tasks={tasks}
                  project={selectedProject}
                  zoomLevel={zoomLevel}
                  settings={ganttSettings}
                  scrollToToday={scrollToToday}
                  onEditTask={this._handleEditTask}
                  onDeleteTask={this._handleDeleteTask}
                  onTaskUpdate={this._handleTaskUpdate}
                  onAddTask={this._handleAddTask}
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
                  tasks={tasks}
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
                  tasks={tasks}
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
                  tasks={tasks}
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
