import * as React from 'react';
import { Callout, DirectionalHint } from '@fluentui/react';
import { IProject, ViewMode, ZoomLevel, IGanttDisplaySettings, ITaskFilter } from '../../models';
import { FilterBar } from './FilterBar';
import styles from './Toolbar.module.scss';

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

export const Toolbar: React.FC<IToolbarProps> = ({
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
    { id: 'dashboard', label: 'Dashboard', icon: '◫' },
    { id: 'list',      label: 'List',      icon: '☰' },
    { id: 'gantt',     label: 'Gantt',     icon: '▬' },
    { id: 'kanban',    label: 'Kanban',    icon: '⬜' },
  ];

  const zooms: { id: ZoomLevel; label: string }[] = [
    { id: 'day',     label: 'Day' },
    { id: 'week',    label: 'Week' },
    { id: 'month',   label: 'Month' },
    { id: 'quarter', label: 'Quarter' },
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
            aria-label="Select project"
          >
            {viewMode === 'portfolio' ? (
              <span style={{ fontSize: 14, marginRight: 2 }}>⊞</span>
            ) : selectedProject ? (
              <div className={styles.projectDot} style={{ background: selectedProject.color }} />
            ) : null}
            <span className={styles.projectName}>
              {viewMode === 'portfolio' ? 'Portfolio' : selectedProject ? selectedProject.title : 'Select a Project'}
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
                <div
                  className={`${styles.calloutItem} ${viewMode === 'portfolio' ? styles.selected : ''}`}
                  onClick={() => { onViewChange('portfolio'); setProjectCalloutVisible(false); }}
                >
                  <span style={{ fontSize: 14, marginRight: 6, flexShrink: 0 }}>⊞</span>
                  <span style={{ flex: 1 }}>Portfolio</span>
                  <span style={{ fontSize: 11, color: '#605E5C' }}>All projects</span>
                </div>
                <div className={styles.calloutSeparator} />
                {projects.map(p => (
                  <div
                    key={p.id}
                    className={`${styles.calloutItem} ${viewMode !== 'portfolio' && selectedProject?.id === p.id ? styles.selected : ''}`}
                    style={p.isArchived ? { opacity: 0.55 } : undefined}
                    onClick={() => { onSelectProject(p); setProjectCalloutVisible(false); }}
                  >
                    <div className={styles.calloutDot} style={{ background: p.color }} />
                    <span style={{ flex: 1 }}>{p.title}</span>
                    {p.isArchived
                      ? <span style={{ fontSize: 10, color: '#605E5C', background: '#F3F2F1', border: '1px solid #EDEBE9', borderRadius: 3, padding: '1px 5px' }}>Archived</span>
                      : <span style={{ fontSize: 11, color: '#605E5C' }}>{p.status}</span>
                    }
                  </div>
                ))}
                {projects.length === 0 && (
                  <div style={{ padding: '10px 16px', color: '#605E5C', fontSize: 13 }}>No projects yet</div>
                )}
                <div className={styles.calloutSeparator} />
                {hasArchivedProjects && (
                  <div
                    className={styles.calloutItem}
                    onClick={() => { onToggleShowArchived(); }}
                    style={{ color: '#605E5C' }}
                  >
                    <span style={{ marginRight: 6, fontSize: 13 }}>{showArchivedProjects ? '☑' : '☐'}</span>
                    <span>Show archived projects</span>
                  </div>
                )}
                <div className={styles.calloutItem} onClick={() => { setProjectCalloutVisible(false); onAddProject(); }}>
                  + New Project
                </div>
                <div className={styles.calloutItem} onClick={() => { setProjectCalloutVisible(false); onImportAsProject(); }}>
                  📥&ensp;Import File as New Project…
                </div>
              </div>
            </Callout>
          )}

          <div className={styles.divider} />

          {selectedProject && viewMode !== 'portfolio' ? (
            <>
              <button className={`${styles.actionBtn} ${styles.primary}`} onClick={onAddTask}>
                + Add Task
              </button>
              <button className={`${styles.actionBtn} ${styles.secondary}`} onClick={onEditProject}>
                Edit Project
              </button>
            </>
          ) : viewMode !== 'portfolio' ? (
            <button className={`${styles.actionBtn} ${styles.primary}`} onClick={onAddProject}>
              + New Project
            </button>
          ) : null}
        </div>

        <div className={styles.row1Right}>
          {selectedProject && viewMode !== 'portfolio' && (
            <button
              className={`${styles.settingsBtn} ${showSettings ? styles.active : ''}`}
              onClick={onOpenSettings}
              title="Options"
            >
              ⚙ Options
            </button>
          )}

          {viewMode === 'portfolio' && (
            <>
              <div ref={moreBtnRef}>
                <button
                  className={styles.iconBtn}
                  onClick={() => setMoreCalloutVisible(v => !v)}
                  title="More options"
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
                    <div className={styles.calloutItem} onClick={() => { setMoreCalloutVisible(false); onPortfolioExportExcel(); }}>
                      📊&ensp;Export to Excel
                    </div>
                    <div className={styles.calloutItem} onClick={() => { setMoreCalloutVisible(false); onPortfolioExportPowerPoint(); }}>
                      📑&ensp;Export to PowerPoint
                    </div>
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
              <button className={styles.todayBtn} onClick={onScrollToToday} title="Scroll to today">
                ◉ Today
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
                  title="More options"
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
                    <div className={styles.calloutItem} onClick={() => { setMoreCalloutVisible(false); onImport(); }}>
                      📥&ensp;Import Tasks…
                    </div>
                    <div className={styles.calloutItem} onClick={() => { setMoreCalloutVisible(false); onExportExcel(); }}>
                      📊&ensp;Export to Excel
                    </div>
                    <div className={styles.calloutItem} onClick={() => { setMoreCalloutVisible(false); onExportPowerPoint(); }}>
                      📑&ensp;Export to PowerPoint
                    </div>
                    {viewMode === 'gantt' && (
                      <div className={styles.calloutItem} onClick={() => { setMoreCalloutVisible(false); onExportImage(); }}>
                        🖼&ensp;Export as Image (PNG)
                      </div>
                    )}
                    <div className={styles.calloutSeparator} />
                    <div className={styles.calloutItem} onClick={() => { setMoreCalloutVisible(false); onEditProject(); }}>
                      ✏️&ensp;Edit Project
                    </div>
                    {selectedProject?.isArchived ? (
                      <div className={styles.calloutItem} onClick={() => { setMoreCalloutVisible(false); onUnarchiveProject(); }}>
                        📂&ensp;Unarchive Project
                      </div>
                    ) : (
                      <div className={styles.calloutItem} onClick={() => { setMoreCalloutVisible(false); onArchiveProject(); }}>
                        🗄️&ensp;Archive Project
                      </div>
                    )}
                    <div className={`${styles.calloutItem} ${styles.danger}`} onClick={() => { setMoreCalloutVisible(false); onDeleteProject(); }}>
                      🗑️&ensp;Send to Recycle Bin
                    </div>
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

export default Toolbar;
