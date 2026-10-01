import * as React from 'react';
import * as strings from 'SmartGanttWebPartStrings';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  Spinner, SpinnerSize, Stack, Dialog, DialogType, DialogFooter,
  DefaultButton, PrimaryButton, MessageBar, MessageBarType, MessageBarButton,
} from '@fluentui/react';

import { SharePointService, isConflictError } from '../services/SharePointService';
import {
  exportTasksToExcel, renderGanttSVG, downloadPNG, exportToPowerPoint,
  exportPortfolioToExcel, exportPortfolioToPowerPoint, exportTasksCsv, exportMilestonesIcs,
} from '../services/ExportService';
import {
  IProject, ITask, IProjectTaskStats, ViewMode, ZoomLevel, TaskStatus,
  IGanttDisplaySettings, DEFAULT_GANTT_SETTINGS, ITaskFilter, EMPTY_TASK_FILTER, isFilterActive,
  IWorkingCalendar, DEFAULT_WORKING_CALENDAR, TASK_STATUS_OPTIONS, TASK_PRIORITY_OPTIONS, DueFilter,
} from '../models';
import { filterTasks } from '../utils/filterUtils';
import { computeShifts, findDependencyCycle } from '../utils/scheduleUtils';
import { PortfolioView } from './views/PortfolioView';
import { Toolbar } from './toolbar/Toolbar';
import { GanttChart } from './gantt/GanttChart';
import { ITaskUpdateOptions } from './gantt/ganttUtils';
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
  /** View shown until the user picks one (their choice is remembered afterwards). */
  defaultView?: ViewMode;
  /** Gantt zoom shown until the user picks one. */
  defaultZoom?: ZoomLevel;
  /** Working days + holidays (web part property pane). Defaults to Mon–Fri. */
  workingCalendar?: IWorkingCalendar;
  /** Configured by the web part; consumed by the service layer, not by this component. */
  registryListName?: string;
}

// User preferences persisted per web part instance (localStorage).
interface IPersistedPrefs {
  viewMode?: ViewMode;
  zoomLevel?: ZoomLevel;
  ganttSettings?: Partial<IGanttDisplaySettings>;
  selectedProjectId?: number | null;
  showArchivedProjects?: boolean;
}

// What a shared "Copy link" URL carries (in the #sg= hash).
interface IShareState {
  view?: ViewMode;
  zoom?: ZoomLevel;
  projectId?: number;
  filter?: Partial<ITaskFilter>;
}

// One undoable edit: the fields a set of tasks had before and after.
interface IHistoryChange { id: number; before: Partial<ITask>; after: Partial<ITask> }
interface IHistoryEntry { projectId: number; changes: IHistoryChange[] }

const HISTORY_LIMIT = 50;
const TOAST_MS = 7000;
const VIEW_MODES: ViewMode[] = ['gantt', 'list', 'kanban', 'dashboard', 'portfolio'];
const ZOOM_LEVELS: ZoomLevel[] = ['day', 'week', 'month', 'quarter'];
const DUE_FILTERS: DueFilter[] = ['all', 'overdue', 'today', 'week'];

interface IToast { message: string; showUndo: boolean }

interface ISmartGanttState {
  projects: IProject[];
  selectedProject: IProject | null;
  tasks: ITask[];
  viewMode: ViewMode;
  zoomLevel: ZoomLevel;
  /** Custom day width from fit-to-project / Ctrl+wheel; null = use zoomLevel. */
  dayWidthOverride: number | null;
  fitSignal: number;
  loading: boolean;
  tasksLoading: boolean;
  error: string | null;
  tasksError: string | null;
  saveError: string | null;
  toast: IToast | null;
  canUndo: boolean;
  canRedo: boolean;
  showProjectPanel: boolean;
  editingProject: IProject | null;
  showTaskPanel: boolean;
  editingTask: ITask | null;
  /** Initial values for a NEW task: a Kanban column's status, or the dates of a Gantt row drag. */
  newTaskDefaults: Partial<ITask> | null;
  deleteProjectConfirm: boolean;
  deleteTaskConfirm: ITask | null;
  bulkDeleteConfirm: number[] | null;
  archiveProjectConfirm: boolean;
  baselineConfirm: boolean;
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
  private _rootRef = React.createRef<HTMLDivElement>();
  // Project id requested by a shared link; consumed by the first project load.
  private _sharedProjectId: number | undefined;
  private _toastTimer: ReturnType<typeof setTimeout> | null = null;

  // Undo/redo for optimistic task edits. Instance fields (not state) since only
  // the two "can undo/redo" flags affect rendering.
  private _undoStack: IHistoryEntry[] = [];
  private _redoStack: IHistoryEntry[] = [];
  // Saves are serialized per task: each save reads the etag left by the
  // previous one, so a rapid second edit doesn't send a stale token and 412.
  private _saveQueues = new Map<number, Promise<unknown>>();

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

  private _calendar(): IWorkingCalendar {
    return this.props.workingCalendar || DEFAULT_WORKING_CALENDAR;
  }

  constructor(props: ISmartGanttProps) {
    super(props);
    this._prefsKey = `SmartGantt_prefs_${props.context.instanceId}`;
    const prefs = this._loadPrefs();
    // A shared link wins over remembered preferences and web part defaults.
    const shared = this._readShareHash();
    this._sharedProjectId = shared?.projectId;
    this.state = {
      projects: [],
      selectedProject: null,
      tasks: [],
      viewMode: shared?.view || prefs.viewMode || props.defaultView || 'gantt',
      zoomLevel: shared?.zoom || prefs.zoomLevel || props.defaultZoom || 'week',
      dayWidthOverride: null,
      fitSignal: 0,
      loading: true,
      tasksLoading: false,
      error: null,
      tasksError: null,
      saveError: null,
      toast: null,
      canUndo: false,
      canRedo: false,
      showProjectPanel: false,
      editingProject: null,
      showTaskPanel: false,
      editingTask: null,
      newTaskDefaults: null,
      deleteProjectConfirm: false,
      deleteTaskConfirm: null,
      bulkDeleteConfirm: null,
      archiveProjectConfirm: false,
      baselineConfirm: false,
      showArchivedProjects: prefs.showArchivedProjects || false,
      scrollToToday: false,
      showImportPanel: false,
      showImportAsProject: false,
      showGanttSettings: false,
      ganttSettings: { ...DEFAULT_GANTT_SETTINGS, ...prefs.ganttSettings },
      taskFilter: shared?.filter ? { ...EMPTY_TASK_FILTER, ...shared.filter } : EMPTY_TASK_FILTER,
      portfolioStats: null,
      portfolioLoading: false,
    };
  }

  private _scrollToTodayTimer: ReturnType<typeof setTimeout> | null = null;
  private _isMounted = false;

  public async componentDidMount(): Promise<void> {
    this._isMounted = true;
    window.addEventListener('keydown', this._handleGlobalKeyDown);
    await this._loadProjects();
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
    window.removeEventListener('keydown', this._handleGlobalKeyDown);
    if (this._scrollToTodayTimer !== null) clearTimeout(this._scrollToTodayTimer);
    if (this._toastTimer !== null) clearTimeout(this._toastTimer);
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

    // The task filter belongs to the project it was built for. Whatever path
    // changed the selection (picker, delete, archive, import-as-new, a reload
    // that fell back to another project) a carried-over filter could hide
    // every task in the new project with no visible cue, so it is dropped
    // here rather than at each call site. The very first selection
    // (prev = none) is exempt so a filter from a shared link survives load.
    const prevId = prevState.selectedProject?.id;
    if (prevId !== undefined && prevId !== s.selectedProject?.id) {
      this._clearHistory();
      if (isFilterActive(s.taskFilter)) this.setState({ taskFilter: EMPTY_TASK_FILTER });
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

  // ─── Shareable link (#sg=…) ──────────────────────────────────────────────

  private _readShareHash(): IShareState | null {
    try {
      const m = /[#&]sg=([^&]*)/.exec(window.location.hash || '');
      if (!m) return null;
      const o = JSON.parse(decodeURIComponent(m[1]));
      if (!o || typeof o !== 'object') return null;
      const out: IShareState = {};
      if (VIEW_MODES.indexOf(o.v) !== -1) out.view = o.v;
      if (ZOOM_LEVELS.indexOf(o.z) !== -1) out.zoom = o.z;
      if (typeof o.p === 'number' && isFinite(o.p)) out.projectId = o.p;
      if (o.f && typeof o.f === 'object') {
        const strs = (a: unknown): string[] => (Array.isArray(a) ? a.filter(x => typeof x === 'string') : []);
        const f: Partial<ITaskFilter> = {};
        if (typeof o.f.t === 'string') f.text = o.f.t;
        f.statuses = strs(o.f.s).filter(x => TASK_STATUS_OPTIONS.indexOf(x as TaskStatus) !== -1) as TaskStatus[];
        f.priorities = strs(o.f.pr).filter(x => TASK_PRIORITY_OPTIONS.indexOf(x as never) !== -1) as ITaskFilter['priorities'];
        f.assignees = strs(o.f.a);
        f.phases = strs(o.f.ph);
        if (DUE_FILTERS.indexOf(o.f.d) !== -1) f.due = o.f.d;
        out.filter = f;
      }
      return out;
    } catch {
      return null;
    }
  }

  private _buildShareUrl(): string {
    const { viewMode, zoomLevel, selectedProject, taskFilter } = this.state;
    const o: Record<string, unknown> = { v: viewMode, z: zoomLevel };
    if (selectedProject) o.p = selectedProject.id;
    if (isFilterActive(taskFilter)) {
      const f: Record<string, unknown> = {};
      if (taskFilter.text) f.t = taskFilter.text;
      if (taskFilter.statuses.length) f.s = taskFilter.statuses;
      if (taskFilter.priorities.length) f.pr = taskFilter.priorities;
      if (taskFilter.assignees.length) f.a = taskFilter.assignees;
      if (taskFilter.phases.length) f.ph = taskFilter.phases;
      if (taskFilter.due !== 'all') f.d = taskFilter.due;
      o.f = f;
    }
    const base = window.location.href.split('#')[0];
    return `${base}#sg=${encodeURIComponent(JSON.stringify(o))}`;
  }

  private _handleCopyLink = async (): Promise<void> => {
    const url = this._buildShareUrl();
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(url);
      } else {
        // Older/locked-down browsers: fall back to a temporary textarea.
        const ta = document.createElement('textarea');
        ta.value = url;
        ta.style.position = 'fixed';
        ta.style.opacity = '0';
        document.body.appendChild(ta);
        ta.select();
        const ok = document.execCommand('copy');
        document.body.removeChild(ta);
        if (!ok) throw new Error(url);
      }
      this._showToast(strings.App_LinkCopied, false);
    } catch (err) {
      this.setState({
        saveError: formatString(strings.App_LinkCopyFailed, { message: this._errMessage(err, 'unknown error') }),
      });
    }
  };

  // ─── Toast / undo history ────────────────────────────────────────────────

  private _showToast(message: string, showUndo: boolean): void {
    if (this._toastTimer !== null) clearTimeout(this._toastTimer);
    this.setState({ toast: { message, showUndo } });
    this._toastTimer = setTimeout(() => {
      this._toastTimer = null;
      if (this._isMounted) this.setState({ toast: null });
    }, TOAST_MS);
  }

  private _dismissToast = (): void => {
    if (this._toastTimer !== null) clearTimeout(this._toastTimer);
    this._toastTimer = null;
    this.setState({ toast: null });
  };

  private _syncHistoryFlags(): void {
    const canUndo = this._undoStack.length > 0;
    const canRedo = this._redoStack.length > 0;
    if (canUndo !== this.state.canUndo || canRedo !== this.state.canRedo) {
      this.setState({ canUndo, canRedo });
    }
  }

  private _clearHistory(): void {
    this._undoStack = [];
    this._redoStack = [];
    this._syncHistoryFlags();
  }

  private _pushHistory(entry: IHistoryEntry): void {
    this._undoStack.push(entry);
    if (this._undoStack.length > HISTORY_LIMIT) this._undoStack.shift();
    this._redoStack = [];
    this._syncHistoryFlags();
  }

  private _undo = (): void => {
    const entry = this._undoStack.pop();
    if (!entry) return;
    if (entry.projectId !== this.state.selectedProject?.id) { this._clearHistory(); return; }
    this._redoStack.push(entry);
    this._syncHistoryFlags();
    void this._applyChanges(entry.changes.map(c => ({ id: c.id, updates: c.before })));
    this._showToast(strings.App_Undone, false);
  };

  private _redo = (): void => {
    const entry = this._redoStack.pop();
    if (!entry) return;
    if (entry.projectId !== this.state.selectedProject?.id) { this._clearHistory(); return; }
    this._undoStack.push(entry);
    this._syncHistoryFlags();
    void this._applyChanges(entry.changes.map(c => ({ id: c.id, updates: c.after })));
    this._showToast(strings.App_Redone, false);
  };

  private _handleGlobalKeyDown = (e: KeyboardEvent): void => {
    if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
    const k = (e.key || '').toLowerCase();
    const isUndo = k === 'z' && !e.shiftKey;
    const isRedo = k === 'y' || (k === 'z' && e.shiftKey);
    if (!isUndo && !isRedo) return;
    const s = this.state;
    if (
      s.viewMode === 'portfolio' || !s.selectedProject ||
      s.showTaskPanel || s.showProjectPanel || s.showImportPanel || s.showImportAsProject ||
      s.deleteTaskConfirm || s.deleteProjectConfirm || s.archiveProjectConfirm ||
      s.bulkDeleteConfirm || s.baselineConfirm
    ) return;
    const t = e.target as HTMLElement | null;
    // Never steal Ctrl+Z from a text field (native undo) or from another web part.
    if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || t.isContentEditable)) return;
    if (t && t !== document.body && this._rootRef.current && !this._rootRef.current.contains(t)) return;
    e.preventDefault();
    if (isUndo) this._undo(); else this._redo();
  };

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
      // Keep the current (or shared-link, or persisted) selection when it
      // still exists, instead of always snapping back to the first project.
      const sharedId = this._sharedProjectId;
      this._sharedProjectId = undefined;
      const wantedId = preferredId
        ?? this.state.selectedProject?.id
        ?? sharedId
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
    // The settings panel only exists for project views; don't leave it flagged
    // open behind the portfolio (it would pop back open on return).
    if (viewMode === 'portfolio') this.setState({ viewMode, showGanttSettings: false });
    else this.setState({ viewMode });
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
    // Picking a preset ends any custom (fit / Ctrl+wheel) zoom.
    this.setState({ zoomLevel, dayWidthOverride: null });
  };

  private _handleDayWidthChange = (dayWidthOverride: number | null): void => {
    this.setState({ dayWidthOverride });
  };

  private _handleFit = (): void => {
    this.setState(s => ({ fitSignal: s.fitSignal + 1 }));
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

  // Bound to a button's onClick, so it must ignore the click event argument.
  private _handleAddTask = (): void => {
    this.setState({ showTaskPanel: true, editingTask: null, newTaskDefaults: null });
  };

  // Kanban column "+ Add Task": the new task starts in that column's status.
  private _handleAddTaskWithStatus = (status?: TaskStatus): void => {
    this.setState({
      showTaskPanel: true,
      editingTask: null,
      newTaskDefaults: typeof status === 'string' ? { status } : null,
    });
  };

  private _handleEditTask = (task: ITask): void => {
    this.setState({ showTaskPanel: true, editingTask: task, newTaskDefaults: null });
  };

  /** Save queue: chains onto any in-flight save of the same task. */
  private _enqueueSave<T>(taskId: number, run: () => Promise<T>): Promise<T> {
    const prev = this._saveQueues.get(taskId) || Promise.resolve();
    const next = prev.then(() => undefined, () => undefined).then(run);
    this._saveQueues.set(taskId, next);
    const cleanup = (): void => {
      if (this._saveQueues.get(taskId) === next) this._saveQueues.delete(taskId);
    };
    next.then(cleanup, cleanup);
    return next;
  }

  private _handleTaskSave = async (data: Partial<ITask>): Promise<void> => {
    const { selectedProject, editingTask } = this.state;
    if (!selectedProject) return;
    try {
      if (editingTask) {
        const taskId = editingTask.id;
        const listName = selectedProject.listName;
        await this._enqueueSave(taskId, () => {
          // Read the token at send time: a drag that finished while the panel
          // was open leaves a newer one in state than the panel's snapshot.
          const current = this.state.tasks.find(t => t.id === taskId);
          return this.props.spService.updateTask(listName, taskId, {
            ...data,
            etag: current?.etag ?? editingTask.etag,
          });
        });
      } else {
        await this.props.spService.createTask(selectedProject.listName, {
          ...data,
          sortOrder: this.state.tasks.length,
        });
      }
      this._clearHistory();
      await this._loadTasks(selectedProject.listName);
      this.setState({ showTaskPanel: false, editingTask: null, newTaskDefaults: null, portfolioStats: null });
    } catch (err) {
      const conflict = isConflictError(err);
      const message = conflict
        ? strings.App_TaskConflict
        : formatString(strings.SmartGantt_CouldNotSaveTask, { message: this._errMessage(err, 'unknown error') });
      this.setState({ saveError: message });
      // Pull the latest tasks (and etags) so the user's next attempt can succeed.
      if (conflict) void this._loadTasks(selectedProject.listName);
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
      this._clearHistory();
      await this._loadTasks(selectedProject.listName);
      this.setState({ portfolioStats: null });
    } catch (err) {
      this.setState({
        saveError: formatString(strings.SmartGantt_CouldNotDeleteTask, { message: this._errMessage(err, 'unknown error') }),
      });
    }
  };

  // List view multi-select delete: the view asks, we confirm and delete.
  private _handleBulkDelete = (ids: number[]): void => {
    if (ids.length > 0) this.setState({ bulkDeleteConfirm: ids });
  };

  private _confirmBulkDelete = async (): Promise<void> => {
    const { selectedProject, bulkDeleteConfirm } = this.state;
    if (!selectedProject || !bulkDeleteConfirm) return;
    this.setState({ bulkDeleteConfirm: null });
    try {
      // Sequential: deleteTask promotes sub-tasks first, so parallel deletes of
      // a parent and its child would race.
      for (let i = 0; i < bulkDeleteConfirm.length; i++) {
        await this.props.spService.deleteTask(selectedProject.listName, bulkDeleteConfirm[i]);
      }
      this._clearHistory();
      await this._loadTasks(selectedProject.listName);
      this.setState({ portfolioStats: null });
    } catch (err) {
      // Some tasks may have been deleted before the failure — show the truth.
      await this._loadTasks(selectedProject.listName);
      this.setState({
        saveError: formatString(strings.SmartGantt_CouldNotDeleteTask, { message: this._errMessage(err, 'unknown error') }),
        portfolioStats: null,
      });
    }
  };

  // Drag across an empty Gantt row: open the New Task panel with those dates
  // filled in (nothing is created until the user saves).
  private _handleCreateTaskAt = (startDate: string, dueDate: string): void => {
    this.setState({
      showTaskPanel: true,
      editingTask: null,
      newTaskDefaults: { startDate, dueDate },
    });
  };

  // Drag from a bar's connector onto another bar: predecessor → successor (finish-to-start).
  private _handleAddDependency = (predId: number, succId: number): void => {
    const { tasks } = this.state;
    const succ = tasks.find(t => t.id === succId);
    if (!succ || predId === succId) return;
    if (succ.dependencies.indexOf(predId) !== -1) {
      this._showToast(strings.App_DependencyExists, false);
      return;
    }
    if (findDependencyCycle(tasks, predId, succId)) {
      this.setState({ saveError: strings.App_DependencyCycle });
      return;
    }
    void this._handleTaskUpdate(succId, {
      dependencies: [...succ.dependencies, predId],
      dependencyLinks: { ...(succ.dependencyLinks || {}), [predId]: { type: 'FS', lag: 0 } },
    }, { cascadeFrom: predId });
  };

  // ─── Exports / print ─────────────────────────────────────────────────────

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

  // Print / save as PDF: render the (filtered) chart into a clean window and
  // hand it to the browser's print dialog, which offers "Save as PDF".
  private _handlePrint = (): void => {
    const { selectedProject, ganttSettings, taskFilter, tasks } = this.state;
    if (!selectedProject) return;
    try {
      const svg = renderGanttSVG(selectedProject, this._getFilteredTasks(tasks, taskFilter), ganttSettings);
      const w = window.open('', '_blank');
      if (!w) {
        this.setState({ saveError: strings.App_PopupBlocked });
        return;
      }
      w.document.title = `${selectedProject.title} - Gantt Chart`;
      const style = w.document.createElement('style');
      style.textContent = '@page { size: landscape; margin: 10mm; } body { margin: 0; font-family: "Segoe UI", sans-serif; } svg { max-width: 100%; height: auto; }';
      w.document.head.appendChild(style);
      w.document.body.innerHTML = svg;
      w.focus();
      // Give the SVG a beat to lay out before the dialog snapshots it.
      setTimeout(() => { w.print(); }, 300);
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

  private _handleExportCsv = (): void => {
    const { selectedProject, tasks } = this.state;
    if (!selectedProject) return;
    try {
      exportTasksCsv(selectedProject, tasks);
    } catch (err) {
      this.setState({
        saveError: formatString(strings.App_CsvExportFailed, { message: this._errMessage(err, 'unknown error') }),
      });
    }
  };

  private _handleExportIcs = (): void => {
    const { selectedProject, tasks } = this.state;
    if (!selectedProject) return;
    try {
      exportMilestonesIcs(selectedProject, tasks);
    } catch (err) {
      this.setState({
        saveError: formatString(strings.App_IcsExportFailed, { message: this._errMessage(err, 'unknown error') }),
      });
    }
  };

  private _handleToggleGanttSettings = (): void => {
    this.setState(s => ({ showGanttSettings: !s.showGanttSettings }));
  };

  private _handleDismissGanttSettings = (): void => {
    this.setState({ showGanttSettings: false });
  };

  private _handleGanttSettingsChange = (ganttSettings: IGanttDisplaySettings): void => {
    this.setState({ ganttSettings });
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

  // ─── Baseline ────────────────────────────────────────────────────────────

  private _handleSetBaseline = (): void => {
    this.setState({ baselineConfirm: true });
  };

  private _confirmSetBaseline = async (): Promise<void> => {
    const { selectedProject, tasks } = this.state;
    this.setState({ baselineConfirm: false });
    if (!selectedProject) return;
    try {
      await this.props.spService.captureBaseline(selectedProject.listName, tasks);
      this._clearHistory();
      await this._loadTasks(selectedProject.listName);
      this._showToast(strings.App_BaselineSaved, false);
    } catch (err) {
      this.setState({
        saveError: formatString(strings.App_BaselineFailed, { message: this._errMessage(err, 'unknown error') }),
      });
    }
  };

  // ─── Optimistic task updates ─────────────────────────────────────────────

  /**
   * Applies field changes to one or more tasks: optimistically in state
   * first, then saved per task (serialized, passing and refreshing each
   * task's etag). Any failure reloads the project's tasks so the screen
   * shows what is really stored; a concurrent-edit conflict says so.
   */
  private async _applyChanges(list: Array<{ id: number; updates: Partial<ITask> }>): Promise<void> {
    const { selectedProject } = this.state;
    if (!selectedProject || list.length === 0) return;
    // Task IDs are only unique within a project's list — capture which
    // project this update belongs to so a project switch mid-flight can't
    // patch/reload the wrong project's tasks.
    const projectId = selectedProject.id;
    const listName = selectedProject.listName;
    const byId = new Map(list.map(c => [c.id, c.updates] as [number, Partial<ITask>]));

    // Optimistic update. Also marks portfolio stats stale so the Portfolio
    // view refetches instead of showing numbers from before this edit.
    this.setState(prev => {
      if (prev.selectedProject?.id !== projectId) return null;
      return {
        tasks: prev.tasks.map(t => (byId.has(t.id) ? { ...t, ...byId.get(t.id) } : t)),
        portfolioStats: null,
      };
    });

    const failures: unknown[] = [];
    await Promise.all(list.map(c =>
      this._enqueueSave(c.id, async () => {
        // Read the etag at send time — the previous save of this task may have
        // just replaced it.
        const cur = this.state.selectedProject?.id === projectId
          ? this.state.tasks.find(t => t.id === c.id)
          : undefined;
        const newEtag = await this.props.spService.updateTask(listName, c.id, { ...c.updates, etag: cur?.etag });
        if (newEtag) {
          this.setState(prev => {
            if (prev.selectedProject?.id !== projectId) return null;
            return { tasks: prev.tasks.map(t => (t.id === c.id ? { ...t, etag: newEtag } : t)) };
          });
        }
      }).catch(err => { failures.push(err); })
    ));

    if (failures.length === 0 || !this._isMounted) return;
    if (this.state.selectedProject?.id !== projectId) return;
    const first = failures[0];
    // Revert to what is actually stored (also fetches fresh etags after a 412).
    this.setState({
      saveError: isConflictError(first)
        ? strings.App_TaskConflict
        : formatString(strings.SmartGantt_CouldNotUpdateTask, { message: this._errMessage(first, 'unknown error') }),
    });
    this._clearHistory();
    await this._loadTasks(listName);
  }

  private _handleTaskUpdate = async (
    taskId: number,
    updates: Partial<ITask>,
    opts: ITaskUpdateOptions = {}
  ): Promise<void> => {
    const { selectedProject, tasks } = this.state;
    if (!selectedProject) return;
    const task = tasks.find(t => t.id === taskId);
    if (!task) return;

    const primary: Partial<ITask> = { ...updates };
    const extras: Array<{ id: number; updates: Partial<ITask> }> = [];

    // Auto-shift dependents: after a date change (or a new link) move every
    // dependent that would now break its link, keeping their durations.
    const cascadeFrom = opts.cascadeFrom !== undefined
      ? opts.cascadeFrom
      : (opts.cascade && (updates.startDate !== undefined || updates.dueDate !== undefined) ? taskId : undefined);
    if (cascadeFrom !== undefined) {
      const merged = tasks.map(t => (t.id === taskId ? { ...t, ...updates } : t));
      const shifts = computeShifts(merged, cascadeFrom, this._calendar());
      shifts.forEach(sh => {
        const dates = { startDate: sh.startDate, dueDate: sh.dueDate };
        if (sh.id === taskId) Object.assign(primary, dates);
        else extras.push({ id: sh.id, updates: dates });
      });
    }

    const list = [{ id: taskId, updates: primary }, ...extras];
    const changes: IHistoryChange[] = list.map(c => {
      const t = tasks.find(x => x.id === c.id)!;
      const before: Partial<ITask> = {};
      (Object.keys(c.updates) as Array<keyof ITask>).forEach(k => {
        if (k === 'etag') return;
        // A missing link map must be restored as "empty", not skipped.
        (before as Record<string, unknown>)[k] = k === 'dependencyLinks' ? (t.dependencyLinks || {}) : t[k];
      });
      return { id: c.id, before, after: c.updates };
    });
    this._pushHistory({ projectId: selectedProject.id, changes });

    if (extras.length > 0) {
      this._showToast(
        extras.length === 1
          ? strings.App_DependentsMovedOne
          : formatString(strings.App_DependentsMovedMany, { count: extras.length }),
        true
      );
    }
    await this._applyChanges(list);
  };

  public render(): React.ReactElement {
    const {
      projects, selectedProject, tasks, viewMode, zoomLevel, dayWidthOverride, fitSignal,
      loading, tasksLoading, error, tasksError, saveError, toast, canUndo, canRedo,
      showProjectPanel, editingProject,
      showTaskPanel, editingTask, newTaskDefaults,
      deleteProjectConfirm, deleteTaskConfirm, bulkDeleteConfirm, archiveProjectConfirm, baselineConfirm,
      showArchivedProjects,
      scrollToToday, showImportPanel, showImportAsProject, showGanttSettings, ganttSettings,
      taskFilter, portfolioStats, portfolioLoading,
    } = this.state;

    const visibleProjects = this._getVisibleProjects(projects, showArchivedProjects);
    const hasArchivedProjects = projects.some(p => p.isArchived);
    const calendar = this._calendar();

    // Derive autocomplete/filter lists from the full task set
    const { phases: knownPhases, users: knownUsers } = this._getKnownLists(tasks);

    const filteredTasks = this._getFilteredTasks(tasks, taskFilter);

    return (
      <div className={styles.smartGantt} ref={this._rootRef}>
        {this.props.title && <div className={styles.webPartTitle}>{this.props.title}</div>}
        <Toolbar
          projects={visibleProjects}
          selectedProject={selectedProject}
          viewMode={viewMode}
          zoomLevel={dayWidthOverride ? null : zoomLevel}
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
          onExportCsv={this._handleExportCsv}
          onExportIcs={this._handleExportIcs}
          onPrint={this._handlePrint}
          onSetBaseline={this._handleSetBaseline}
          onCopyLink={this._handleCopyLink}
          onFit={this._handleFit}
          onUndo={this._undo}
          onRedo={this._redo}
          canUndo={canUndo}
          canRedo={canRedo}
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

        {toast && (
          <MessageBar
            messageBarType={MessageBarType.info}
            onDismiss={this._dismissToast}
            dismissButtonAriaLabel={strings.SmartGantt_DismissAriaLabel}
            actions={toast.showUndo ? (
              <div>
                <MessageBarButton onClick={() => { this._dismissToast(); this._undo(); }}>
                  {strings.App_UndoButton}
                </MessageBarButton>
              </div>
            ) : undefined}
          >
            {toast.message}
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
                  onClick={this._handleShowImportAsProject}
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
              calendar={calendar}
            />
          )}

          {!loading && !error && selectedProject && viewMode !== 'portfolio' && (
            <>
              {tasksLoading && (
                <div className={styles.loadingContainer} style={{ position: 'absolute', zIndex: 10, background: 'var(--sg-overlay, rgba(255,255,255,0.8))' }}>
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
                  calendar={calendar}
                  dayWidthOverride={dayWidthOverride}
                  onDayWidthChange={this._handleDayWidthChange}
                  fitSignal={fitSignal}
                  onEditTask={this._handleEditTask}
                  onDeleteTask={this._handleDeleteTask}
                  onTaskUpdate={this._handleTaskUpdate}
                  onAddTask={this._handleAddTask}
                  onImport={this._handleShowImportPanel}
                  onCreateTaskAt={this._handleCreateTaskAt}
                  onAddDependency={this._handleAddDependency}
                />
              )}

              {/* GanttSettings sits outside the viewMode blocks (but inside the
                  project block) so the Options button works from every project
                  view: the panel stays mounted and showGanttSettings is always
                  reflected. The button itself is hidden in Portfolio, and
                  switching to Portfolio closes the panel (_handleViewChange). */}
              <GanttSettings
                isOpen={showGanttSettings}
                settings={ganttSettings}
                onChange={this._handleGanttSettingsChange}
                onDismiss={this._handleDismissGanttSettings}
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
                  onBulkDelete={this._handleBulkDelete}
                  calendar={calendar}
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
                  onAddTask={this._handleAddTaskWithStatus}
                />
              )}

              {viewMode === 'dashboard' && (
                <DashboardView
                  tasks={filteredTasks}
                  project={selectedProject}
                  onEditTask={this._handleEditTask}
                  onAddTask={this._handleAddTask}
                  calendar={calendar}
                />
              )}
            </>
          )}
        </div>

        {/* Project Panel */}
        <ProjectPanel
          isOpen={showProjectPanel}
          project={editingProject}
          spService={this.props.spService}
          knownUsers={knownUsers}
          onSave={this._handleProjectSave}
          onDismiss={() => this.setState({ showProjectPanel: false, editingProject: null })}
        />

        {/* Task Panel */}
        {selectedProject && (
          <TaskPanel
            isOpen={showTaskPanel}
            task={editingTask}
            tasks={tasks}
            allTasks={tasks}
            project={selectedProject}
            knownPhases={knownPhases}
            knownUsers={knownUsers}
            calendar={calendar}
            spService={this.props.spService}
            defaults={newTaskDefaults || undefined}
            onSave={this._handleTaskSave}
            onDismiss={() => this.setState({ showTaskPanel: false, editingTask: null, newTaskDefaults: null })}
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
              this.setState({ showImportPanel: false, portfolioStats: null });
              this._clearHistory();
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
            this.setState({ showImportAsProject: false, portfolioStats: null });
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

        {/* Bulk delete tasks confirm (List view multi-select) */}
        <Dialog
          hidden={!bulkDeleteConfirm}
          onDismiss={() => this.setState({ bulkDeleteConfirm: null })}
          dialogContentProps={{
            type: DialogType.normal,
            title: strings.App_BulkDeleteDialogTitle,
            subText: formatString(strings.App_BulkDeleteDialogSubtext, { count: bulkDeleteConfirm ? bulkDeleteConfirm.length : 0 }),
          }}
        >
          <DialogFooter>
            <DefaultButton text={strings.SmartGantt_CancelButton} onClick={() => this.setState({ bulkDeleteConfirm: null })} />
            <PrimaryButton
              text={strings.SmartGantt_DeleteButton}
              styles={{ root: { background: '#D13438', borderColor: '#D13438' } }}
              onClick={this._confirmBulkDelete}
            />
          </DialogFooter>
        </Dialog>

        {/* Set baseline confirm */}
        <Dialog
          hidden={!baselineConfirm}
          onDismiss={() => this.setState({ baselineConfirm: false })}
          dialogContentProps={{
            type: DialogType.normal,
            title: strings.App_SetBaselineDialogTitle,
            subText: formatString(strings.App_SetBaselineDialogSubtext, { count: tasks.length }),
          }}
        >
          <DialogFooter>
            <DefaultButton text={strings.SmartGantt_CancelButton} onClick={() => this.setState({ baselineConfirm: false })} />
            <PrimaryButton text={strings.App_SetBaselineButton} onClick={this._confirmSetBaseline} />
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
