import * as React from 'react';
import { Callout, DirectionalHint } from '@fluentui/react';
import { IProject, ViewMode, ZoomLevel, IGanttDisplaySettings, ITaskFilter } from '../../models';
import { FilterBar } from './FilterBar';
import * as strings from 'SmartGanttWebPartStrings';
import styles from './Toolbar.module.scss';

// Callout items were plain divs with onClick only — reachable by mouse but
// not by keyboard (a Callout can be opened via Enter/Space, but nothing
// inside it could then be activated without a pointer). Centralized here
// since the pattern repeats across the project selector and both "more
// options" menus.
const CalloutMenuItem: React.FC<{
  onClick: () => void;
  className?: string;
  style?: React.CSSProperties;
  children: React.ReactNode;
}> = ({ onClick, className, style, children }) => (
  <div
    className={className}
    style={style}
    role="menuitem"
    tabIndex={0}
    onClick={onClick}
    onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onClick(); } }}
  >
    {children}
  </div>
);

interface IToolbarProps {
  projects: IProject[];
  selectedProject: IProject | null;
  viewMode: ViewMode;
  zoomLevel: ZoomLevel;
  ganttSettings: IGanttDisplaySettings;
  onSelectProject: (project: IProject) => void;
  onViewChange: (view: ViewMode) => void;
  onZoomChange: (zoom: ZoomLevel) => void;
  onScrollToToday: () => void;
  onAddTask: () => void;
  onAddProject: () => void;
  onEditProject: () => void;
  onDeleteProject: () => void;
  onArchiveProject: () => void;
  onUnarchiveProject: () => void;
  showArchivedProjects: boolean;
  hasArchivedProjects: boolean;
  onToggleShowArchived: () => void;
  onImport: () => void;
  onImportAsProject: () => void;
  onExportExcel: () => void;
  onExportImage: () => void;
  onExportPowerPoint: () => void;
  onPortfolioExportExcel: () => void;
  onPortfolioExportPowerPoint: () => void;
  onOpenSettings: () => void;
  showSettings: boolean;
  taskFilter: ITaskFilter;
  onFilterChange: (f: ITaskFilter) => void;
  knownUsers: string[];
  knownPhases: string[];
  filteredCount: number;
  totalCount: number;
}

const ToolbarComponent: React.FC<IToolbarProps> = ({
  projects,
  selectedProject,
  viewMode,
  zoomLevel,
  onSelectProject,
  onViewChange,
  onZoomChange,
  onScrollToToday,
  onAddTask,
  onAddProject,
  onEditProject,
  onDeleteProject,
  onArchiveProject,
  onUnarchiveProject,
  showArchivedProjects,
  hasArchivedProjects,
  onToggleShowArchived,
  onImport,
  onImportAsProject,
  onExportExcel,
  onExportImage,
  onExportPowerPoint,
  onPortfolioExportExcel,
  onPortfolioExportPowerPoint,
  onOpenSettings,
  showSettings,
  taskFilter,
  onFilterChange,
  knownUsers,
  knownPhases,
  filteredCount,
  totalCount,
}) => {
  const [projectCalloutVisible, setProjectCalloutVisible] = React.useState(false);
  const [moreCalloutVisible, setMoreCalloutVisible] = React.useState(false);
  const projectBtnRef = React.useRef<HTMLDivElement>(null);
  const moreBtnRef = React.useRef<HTMLDivElement>(null);

  const views: { id: ViewMode; label: string; icon: string }[] = [
    { id: 'dashboard', label: strings.Toolbar_ViewDashboard, icon: '◫' },
    { id: 'list',      label: strings.Toolbar_ViewList,      icon: '☰' },
    { id: 'gantt',     label: strings.Toolbar_ViewGantt,     icon: '▬' },
    { id: 'kanban',    label: strings.Toolbar_ViewKanban,    icon: '⬜' },
  ];

  const zooms: { id: ZoomLevel; label: string }[] = [
    { id: 'day',     label: strings.Toolbar_ZoomDay },
    { id: 'week',    label: strings.Toolbar_ZoomWeek },
    { id: 'month',   label: strings.Toolbar_ZoomMonth },
    { id: 'quarter', label: strings.Toolbar_ZoomQuarter },
  ];

  return (
    <div className={styles.toolbar}>
      {/* ── Row 1: project selector + task actions ───────────────────────── */}
      <div className={styles.row1}>
        <div className={styles.row1Left}>
          {/* Project selector */}
          <div
            className={styles.projectSelector}
            ref={projectBtnRef}
            onClick={() => setProjectCalloutVisible(v => !v)}
            onKeyDown={e => {
              if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                setProjectCalloutVisible(v => !v);
              }
            }}
            role="button"
            tabIndex={0}
            aria-haspopup="true"
            aria-expanded={projectCalloutVisible}
            aria-label={strings.Toolbar_SelectProjectAriaLabel}
          >
            {viewMode === 'portfolio' ? (
              <span style={{ fontSize: 14, marginRight: 2 }}>⊞</span>
            ) : selectedProject ? (
              <div className={styles.projectDot} style={{ background: selectedProject.color }} />
            ) : null}
            <span className={styles.projectName}>
              {viewMode === 'portfolio' ? strings.Toolbar_PortfolioLabel : selectedProject ? selectedProject.title : strings.Toolbar_SelectAProject}
            </span>
            <span className={styles.chevron}>▾</span>
          </div>

          {projectCalloutVisible && (
            <Callout
              target={projectBtnRef}
              onDismiss={() => setProjectCalloutVisible(false)}
              directionalHint={DirectionalHint.bottomLeftEdge}
              calloutMinWidth={260}
            >
              <div style={{ paddingTop: 4, paddingBottom: 4 }}>
                {/* Portfolio — cross-project overview */}
                <CalloutMenuItem
                  className={`${styles.calloutItem} ${viewMode === 'portfolio' ? styles.selected : ''}`}
                  onClick={() => { onViewChange('portfolio'); setProjectCalloutVisible(false); }}
                >
                  <span style={{ fontSize: 14, marginRight: 6, flexShrink: 0 }}>⊞</span>
                  <span style={{ flex: 1 }}>{strings.Toolbar_PortfolioLabel}</span>
                  <span style={{ fontSize: 11, color: '#605E5C' }}>{strings.Toolbar_AllProjectsSubtext}</span>
                </CalloutMenuItem>
                <div className={styles.calloutSeparator} />
                {projects.map(p => (
                  <CalloutMenuItem
                    key={p.id}
                    className={`${styles.calloutItem} ${viewMode !== 'portfolio' && selectedProject?.id === p.id ? styles.selected : ''}`}
                    style={p.isArchived ? { opacity: 0.55 } : undefined}
                    onClick={() => { onSelectProject(p); setProjectCalloutVisible(false); }}
                  >
                    <div className={styles.calloutDot} style={{ background: p.color }} />
                    <span style={{ flex: 1 }}>{p.title}</span>
                    {p.isArchived
                      ? <span style={{ fontSize: 10, color: '#605E5C', background: '#F3F2F1', border: '1px solid #EDEBE9', borderRadius: 3, padding: '1px 5px' }}>{strings.Toolbar_ArchivedBadge}</span>
                      : <span style={{ fontSize: 11, color: '#605E5C' }}>{p.status}</span>
                    }
                  </CalloutMenuItem>
                ))}
                {projects.length === 0 && (
                  <div style={{ padding: '10px 16px', color: '#605E5C', fontSize: 13 }}>{strings.Toolbar_NoProjectsYet}</div>
                )}
                <div className={styles.calloutSeparator} />
                {hasArchivedProjects && (
                  <CalloutMenuItem
                    className={styles.calloutItem}
                    onClick={() => { onToggleShowArchived(); }}
                    style={{ color: '#605E5C' }}
                  >
                    <span style={{ marginRight: 6, fontSize: 13 }}>{showArchivedProjects ? '☑' : '☐'}</span>
                    <span>{strings.Toolbar_ShowArchivedProjects}</span>
                  </CalloutMenuItem>
                )}
                <CalloutMenuItem className={styles.calloutItem} onClick={() => { setProjectCalloutVisible(false); onAddProject(); }}>
                  {strings.Toolbar_NewProject}
                </CalloutMenuItem>
                <CalloutMenuItem className={styles.calloutItem} onClick={() => { setProjectCalloutVisible(false); onImportAsProject(); }}>
                  {strings.Toolbar_ImportFileAsNewProject}
                </CalloutMenuItem>
              </div>
            </Callout>
          )}

          <div className={styles.divider} />

          {selectedProject && viewMode !== 'portfolio' ? (
            <>
              <button className={`${styles.actionBtn} ${styles.primary}`} onClick={onAddTask}>
                {strings.Toolbar_AddTask}
              </button>
              <button className={`${styles.actionBtn} ${styles.secondary}`} onClick={onEditProject}>
                {strings.Toolbar_EditProject}
              </button>
            </>
          ) : viewMode !== 'portfolio' ? (
            <button className={`${styles.actionBtn} ${styles.primary}`} onClick={onAddProject}>
              {strings.Toolbar_NewProject}
            </button>
          ) : null}
        </div>

        <div className={styles.row1Right}>
          {selectedProject && viewMode !== 'portfolio' && (
            <button
              className={`${styles.settingsBtn} ${showSettings ? styles.active : ''}`}
              onClick={onOpenSettings}
              title={strings.Toolbar_OptionsButton}
            >
              {strings.Toolbar_OptionsButton}
            </button>
          )}

          {viewMode === 'portfolio' && (
            <>
              <div ref={moreBtnRef}>
                <button
                  className={styles.iconBtn}
                  onClick={() => setMoreCalloutVisible(v => !v)}
                  title={strings.Toolbar_MoreOptionsTitle}
                >
                  ⋯
                </button>
              </div>
              {moreCalloutVisible && (
                <Callout
                  target={moreBtnRef}
                  onDismiss={() => setMoreCalloutVisible(false)}
                  directionalHint={DirectionalHint.bottomRightEdge}
                  calloutMinWidth={210}
                >
                  <div style={{ paddingTop: 4, paddingBottom: 4 }}>
                    <CalloutMenuItem className={styles.calloutItem} onClick={() => { setMoreCalloutVisible(false); onPortfolioExportExcel(); }}>
                      {strings.Toolbar_ExportToExcel}
                    </CalloutMenuItem>
                    <CalloutMenuItem className={styles.calloutItem} onClick={() => { setMoreCalloutVisible(false); onPortfolioExportPowerPoint(); }}>
                      {strings.Toolbar_ExportToPowerPoint}
                    </CalloutMenuItem>
                  </div>
                </Callout>
              )}
            </>
          )}
        </div>
      </div>

      {/* ── Row 2: view controls (hidden in portfolio mode) ─────────────── */}
      {viewMode !== 'portfolio' && <div className={styles.row2}>
        <div className={styles.row2Left}>
          {/* Zoom + Today — Gantt view only */}
          {viewMode === 'gantt' && (
            <>
              <button className={styles.todayBtn} onClick={onScrollToToday} title={strings.Toolbar_ScrollToTodayTitle}>
                {strings.Toolbar_TodayButton}
              </button>
              <div className={styles.zoomGroup}>
                {zooms.map(z => (
                  <button
                    key={z.id}
                    className={`${styles.zoomBtn} ${zoomLevel === z.id ? styles.active : ''}`}
                    onClick={() => onZoomChange(z.id)}
                    aria-pressed={zoomLevel === z.id}
                  >
                    {z.label}
                  </button>
                ))}
              </div>
            </>
          )}
        </div>

        <div className={styles.row2Right}>
          {/* View switcher */}
          <div className={styles.viewSwitcher}>
            {views.map(v => (
              <button
                key={v.id}
                className={`${styles.viewBtn} ${viewMode === v.id ? styles.active : ''}`}
                onClick={() => onViewChange(v.id)}
              >
                <span>{v.icon}</span>
                <span>{v.label}</span>
              </button>
            ))}
          </div>

          {/* ⋯ More (project-level actions) */}
          {selectedProject && (
            <>
              <div ref={moreBtnRef}>
                <button
                  className={styles.iconBtn}
                  onClick={() => setMoreCalloutVisible(v => !v)}
                  title={strings.Toolbar_MoreOptionsTitle}
                >
                  ⋯
                </button>
              </div>
              {moreCalloutVisible && (
                <Callout
                  target={moreBtnRef}
                  onDismiss={() => setMoreCalloutVisible(false)}
                  directionalHint={DirectionalHint.bottomRightEdge}
                  calloutMinWidth={200}
                >
                  <div style={{ paddingTop: 4, paddingBottom: 4 }}>
                    <CalloutMenuItem className={styles.calloutItem} onClick={() => { setMoreCalloutVisible(false); onImport(); }}>
                      {strings.Toolbar_ImportTasks}
                    </CalloutMenuItem>
                    <CalloutMenuItem className={styles.calloutItem} onClick={() => { setMoreCalloutVisible(false); onExportExcel(); }}>
                      {strings.Toolbar_ExportToExcel}
                    </CalloutMenuItem>
                    <CalloutMenuItem className={styles.calloutItem} onClick={() => { setMoreCalloutVisible(false); onExportPowerPoint(); }}>
                      {strings.Toolbar_ExportToPowerPoint}
                    </CalloutMenuItem>
                    {viewMode === 'gantt' && (
                      <CalloutMenuItem className={styles.calloutItem} onClick={() => { setMoreCalloutVisible(false); onExportImage(); }}>
                        {strings.Toolbar_ExportAsImage}
                      </CalloutMenuItem>
                    )}
                    <div className={styles.calloutSeparator} />
                    <CalloutMenuItem className={styles.calloutItem} onClick={() => { setMoreCalloutVisible(false); onEditProject(); }}>
                      ✏️&ensp;{strings.Toolbar_EditProject}
                    </CalloutMenuItem>
                    {selectedProject?.isArchived ? (
                      <CalloutMenuItem className={styles.calloutItem} onClick={() => { setMoreCalloutVisible(false); onUnarchiveProject(); }}>
                        {strings.Toolbar_UnarchiveProject}
                      </CalloutMenuItem>
                    ) : (
                      <CalloutMenuItem className={styles.calloutItem} onClick={() => { setMoreCalloutVisible(false); onArchiveProject(); }}>
                        {strings.Toolbar_ArchiveProject}
                      </CalloutMenuItem>
                    )}
                    <CalloutMenuItem className={`${styles.calloutItem} ${styles.danger}`} onClick={() => { setMoreCalloutVisible(false); onDeleteProject(); }}>
                      {strings.Toolbar_SendToRecycleBin}
                    </CalloutMenuItem>
                  </div>
                </Callout>
              )}
            </>
          )}
        </div>
      </div>}

      {/* ── Row 3: search + filters (own row so they have room to breathe) ── */}
      {viewMode !== 'portfolio' && selectedProject && totalCount > 0 && (
        <div className={styles.row3}>
          <FilterBar
            filter={taskFilter}
            onChange={onFilterChange}
            assignees={knownUsers}
            phases={knownPhases}
            matchCount={filteredCount}
            totalCount={totalCount}
          />
        </div>
      )}
    </div>
  );
};

export const Toolbar = React.memo(ToolbarComponent);

export default Toolbar;
