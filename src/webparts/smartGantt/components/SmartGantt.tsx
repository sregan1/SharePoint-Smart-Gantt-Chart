import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Spinner, SpinnerSize, Dialog, DialogType, DialogFooter, DefaultButton, PrimaryButton } from '@fluentui/react';

import { SharePointService } from '../services/SharePointService';
import { exportTasksToExcel, renderGanttSVG, downloadPNG, exportToPowerPoint } from '../services/ExportService';
import { IProject, ITask, ViewMode, ZoomLevel, IGanttDisplaySettings, DEFAULT_GANTT_SETTINGS } from '../models';
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
  scrollToToday: boolean;
  showImportPanel: boolean;
  showGanttSettings: boolean;
  ganttSettings: IGanttDisplaySettings;
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
      scrollToToday: false,
      showImportPanel: false,
      showGanttSettings: false,
      ganttSettings: DEFAULT_GANTT_SETTINGS,
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
      this.setState({ showProjectPanel: false, editingProject: null });
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
      this.setState({
        projects,
        selectedProject: projects[0] || null,
        tasks: [],
        deleteProjectConfirm: false,
      });
      if (projects[0]) {
        await this._loadTasks(projects[0].listName);
      }
    } catch (_err) {
      this.setState({ deleteProjectConfirm: false });
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
      this.setState({ showTaskPanel: false, editingTask: null });
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
      deleteProjectConfirm, scrollToToday,
      showImportPanel, showGanttSettings, ganttSettings,
    } = this.state;

    // Derive autocomplete lists from current tasks
    const knownPhases = Array.from(new Set(tasks.map(t => t.phase).filter(Boolean))).sort() as string[];
    const knownUsers = Array.from(new Set(tasks.map(t => t.assignedTo).filter(Boolean))).sort() as string[];

    return (
      <div className={styles.smartGantt}>
        <Toolbar
          projects={projects}
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
          onImport={() => this.setState({ showImportPanel: true })}
          onExportExcel={() => { if (selectedProject) exportTasksToExcel(selectedProject, tasks); }}
          onExportImage={this._handleExportImage}
          onExportPowerPoint={this._handleExportPowerPoint}
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

          {!loading && !error && !selectedProject && (
            <div className={styles.emptyState}>
              <div className={styles.emptyIcon}>📋</div>
              <div className={styles.emptyTitle}>No projects yet</div>
              <div className={styles.emptySubtitle}>
                Create your first project to start managing tasks with Gantt charts, lists, and Kanban boards.
              </div>
              <PrimaryButton text="+ Create First Project" onClick={this._handleAddProject} />
            </div>
          )}

          {!loading && !error && selectedProject && (
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

        {/* Delete project confirm */}
        <Dialog
          hidden={!deleteProjectConfirm}
          onDismiss={() => this.setState({ deleteProjectConfirm: false })}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Delete Project?',
            subText: `This will permanently delete "${selectedProject?.title}" and all its tasks. This action cannot be undone.`,
          }}
        >
          <DialogFooter>
            <DefaultButton text="Cancel" onClick={() => this.setState({ deleteProjectConfirm: false })} />
            <PrimaryButton
              text="Delete"
              styles={{ root: { background: '#D13438', borderColor: '#D13438' } }}
              onClick={this._confirmDeleteProject}
            />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
}
