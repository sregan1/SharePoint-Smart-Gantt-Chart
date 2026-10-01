import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { differenceInCalendarDays } from 'date-fns';
import * as strings from 'SmartGanttWebPartStrings';
import { formatString } from '../localeUtils';
import { IProject, IProjectTaskStats, IWorkingCalendar, PROJECT_STATUS_COLORS, PROJECT_STATUS_LIGHT_COLORS, PROJECT_STATUS_OPTIONS } from '../../models';
import { HealthBadge } from '../common/HealthBadge';
import { ProjectHealth } from '../../models';
import { parseDateOnly, formatDateOnly, todayLocalMidnight } from '../../utils/dateUtils';
import { getProjectStatusLabel } from '../../utils/taskDisplayUtils';
import { onActivate } from '../common/a11y';
import { readStored, writeStored } from '../common/storage';
import { SHARED_VIEW_CSS } from '../common/theme';

// ─── Types ────────────────────────────────────────────────────────────────────

export interface IPortfolioViewProps {
  projects: IProject[];
  statsMap: Map<number, IProjectTaskStats> | null;
  loading: boolean;
  onSelectProject: (project: IProject) => void;
  onAddProject: () => void;
  onRefresh: () => void;
  /** Accepted for parity with the other views; the portfolio has no date-based
   *  calendar logic of its own, so it is currently unused. */
  calendar?: IWorkingCalendar;
}

type SortKey = 'name' | 'health' | 'status' | 'completion';

const HEALTH_ORDER: Record<ProjectHealth, number> = { overdue: 0, 'at-risk': 1, 'on-track': 2, complete: 3 };

const PREFS_KEY = 'sg.portfolio.prefs';
const SORT_KEYS: SortKey[] = ['name', 'health', 'status', 'completion'];

interface IPortfolioPrefs {
  sortKey: SortKey;
  hideDone: boolean;
  search: string;
}

function formatDate(s: string): string {
  return formatDateOnly(s, 'MMM d, yyyy');
}

function initials(name: string): string {
  return name.split(' ').map(w => w[0]).join('').toUpperCase().slice(0, 2);
}

// ─── Mini timeline ────────────────────────────────────────────────────────────

interface IMiniTimelineProps {
  start: string;
  end: string;
  color: string;
}

const MiniTimeline: React.FC<IMiniTimelineProps> = ({ start, end, color }) => {
  if (!start || !end) return null;

  const startDate = parseDateOnly(start);
  const endDate = parseDateOnly(end);
  const today = todayLocalMidnight();

  if (!startDate || !endDate) return null;

  // A single-day project (start === end) is a legitimate 1-day span, not a
  // reason to render nothing — only a genuinely inverted range bails out.
  if (endDate < startDate) return null;
  const total = Math.max(1, differenceInCalendarDays(endDate, startDate));

  const elapsed = Math.max(0, Math.min(total, differenceInCalendarDays(today, startDate)));
  const todayPct = Math.round((elapsed / total) * 100);
  const clampedPct = Math.max(0, Math.min(100, todayPct));

  return (
    <div style={{ marginTop: 12 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 11, color: 'var(--neutralSecondary, #605E5C)', marginBottom: 4 }}>
        <span>{formatDate(start)}</span>
        <span>{formatDate(end)}</span>
      </div>
      <div style={{ position: 'relative', height: 6, background: 'var(--neutralLight, #EDEBE9)', borderRadius: 3, overflow: 'visible' }}>
        <div className="sgFill" style={{
          position: 'absolute',
          left: 0,
          top: 0,
          height: '100%',
          width: `${clampedPct}%`,
          background: `${color}40`,
          borderRadius: 3,
        }} />
        {clampedPct > 0 && clampedPct < 100 && (
          <div className="sgFill" style={{
            position: 'absolute',
            left: `${clampedPct}%`,
            top: -4,
            width: 2,
            height: 14,
            background: '#D13438',
            borderRadius: 1,
            transform: 'translateX(-50%)',
          }}
            title={strings.Portfolio_TodayTooltip}
          />
        )}
      </div>
    </div>
  );
};

// ─── Project Card ─────────────────────────────────────────────────────────────

interface IProjectCardProps {
  project: IProject;
  stats: IProjectTaskStats | undefined;
  statsLoading: boolean;
  onClick: () => void;
}

const ProjectCard: React.FC<IProjectCardProps> = ({ project, stats, statsLoading, onClick }) => {
  const statusColor = PROJECT_STATUS_COLORS[project.status] || '#8B929A';
  const statusBg = PROJECT_STATUS_LIGHT_COLORS[project.status] || 'var(--neutralLighter, #F3F2F1)';

  return (
    <div
      className="sgCard sgFocusable"
      onClick={onClick}
      role="button"
      tabIndex={0}
      aria-label={formatString(strings.View_Portfolio_CardAriaLabel, { projectName: project.title, status: getProjectStatusLabel(project.status) })}
      onKeyDown={onActivate(onClick)}
      style={{
        background: 'var(--white, #fff)',
        borderRadius: 8,
        boxShadow: '0 1px 4px rgba(0,0,0,0.10)',
        borderLeft: `4px solid ${project.color || '#0078D4'}`,
        padding: '16px 18px',
        cursor: 'pointer',
        transition: 'box-shadow 0.15s, transform 0.1s',
        display: 'flex',
        flexDirection: 'column',
        gap: 0,
      }}
      onMouseEnter={e => {
        (e.currentTarget as HTMLDivElement).style.boxShadow = '0 4px 16px rgba(0,0,0,0.14)';
        (e.currentTarget as HTMLDivElement).style.transform = 'translateY(-1px)';
      }}
      onMouseLeave={e => {
        (e.currentTarget as HTMLDivElement).style.boxShadow = '0 1px 4px rgba(0,0,0,0.10)';
        (e.currentTarget as HTMLDivElement).style.transform = '';
      }}
    >
      {/* Header: title + manager avatar */}
      <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10, marginBottom: 10 }}>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: 15, fontWeight: 600, color: 'var(--neutralPrimary, #1B1B1B)', marginBottom: 4, lineHeight: 1.3 }}>
            {project.title}
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
            <span style={{
              display: 'inline-block',
              padding: '2px 8px',
              borderRadius: 10,
              background: statusBg,
              color: statusColor,
              fontSize: 11,
              fontWeight: 600,
            }}>
              {getProjectStatusLabel(project.status)}
            </span>
            {stats && !stats.statsError && <HealthBadge health={stats.health} size="md" />}
            {stats?.statsError && (
              <span style={{
                display: 'inline-block', padding: '2px 8px', borderRadius: 10,
                background: '#FDF3F4', color: '#D13438', fontSize: 11, fontWeight: 600,
              }}>
                {strings.Portfolio_StatsUnavailable}
              </span>
            )}
          </div>
        </div>
        {project.projectManager && (
          <div
            title={project.projectManager}
            style={{
              width: 32,
              height: 32,
              borderRadius: '50%',
              background: project.color || '#0078D4',
              color: '#fff',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              fontSize: 12,
              fontWeight: 700,
              flexShrink: 0,
            }}
          >
            {initials(project.projectManager)}
          </div>
        )}
      </div>

      {/* Progress bar */}
      {statsLoading && !stats ? (
        <div style={{ height: 32, background: 'var(--neutralLighter, #F3F2F1)', borderRadius: 4, animation: 'pulse 1.5s ease-in-out infinite' }} />
      ) : stats?.statsError ? (
        <div style={{ fontSize: 12, color: 'var(--neutralSecondary, #605E5C)', padding: '8px 0' }}>
          {strings.Portfolio_StatsUnavailableDetail}
        </div>
      ) : stats ? (
        <>
          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, color: 'var(--neutralSecondary, #605E5C)', marginBottom: 4 }}>
            <span>{strings.Portfolio_OverallProgressLabel}</span>
            <span style={{ fontWeight: 600, color: 'var(--neutralPrimary, #1B1B1B)' }}>{stats.overallPct}%</span>
          </div>
          <div style={{ height: 8, background: 'var(--neutralLight, #EDEBE9)', borderRadius: 4, overflow: 'hidden', marginBottom: 12 }}>
            <div className="sgFill" style={{
              height: '100%',
              width: `${stats.overallPct}%`,
              background: project.color || '#0078D4',
              borderRadius: 4,
              transition: 'width 0.4s ease',
            }} />
          </div>

          {/* Task count grid */}
          <div style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(4, 1fr)',
            gap: 6,
            marginBottom: 4,
          }}>
            {([
              { label: strings.Portfolio_TaskCountDone, value: stats.completedCount, color: '#107C10' },
              { label: strings.Portfolio_TaskCountActive, value: stats.inProgressCount, color: 'var(--themePrimary, #0078D4)' },
              { label: strings.Portfolio_TaskCountAtRisk, value: stats.atRiskCount, color: '#CA5010' },
              { label: strings.Portfolio_TaskCountOverdue, value: stats.overdueCount, color: '#D13438' },
            ]).map(({ label, value, color }) => (
              <div key={label} style={{ textAlign: 'center', padding: '6px 4px', background: 'var(--neutralLighterAlt, #FAF9F8)', borderRadius: 6 }}>
                <div style={{ fontSize: 18, fontWeight: 700, color, lineHeight: 1 }}>{value}</div>
                <div style={{ fontSize: 10, color: 'var(--neutralSecondary, #605E5C)', marginTop: 2 }}>{label}</div>
              </div>
            ))}
          </div>

          {/* Mini timeline */}
          <MiniTimeline
            start={stats.earliestStart || project.startDate}
            end={stats.latestDue || project.dueDate}
            color={project.color || '#0078D4'}
          />
        </>
      ) : null}

      {/* Footer link */}
      <div style={{ marginTop: 12, paddingTop: 10, borderTop: '1px solid var(--neutralLighter, #F3F2F1)', display: 'flex', justifyContent: 'flex-end' }}>
        <span style={{ fontSize: 12, color: 'var(--themePrimary, #0078D4)', fontWeight: 600 }}>
          {strings.Portfolio_ViewGanttLink}
        </span>
      </div>
    </div>
  );
};

// ─── Portfolio View ───────────────────────────────────────────────────────────

export const PortfolioView: React.FC<IPortfolioViewProps> = ({
  projects,
  statsMap,
  loading,
  onSelectProject,
  onAddProject,
  onRefresh,
}) => {
  // Sort, search and the hide-finished toggle are remembered per viewer.
  const [prefs, setPrefs] = React.useState<IPortfolioPrefs>(() => {
    const stored = readStored<Partial<IPortfolioPrefs>>(PREFS_KEY, {});
    return {
      sortKey: stored.sortKey && SORT_KEYS.indexOf(stored.sortKey) !== -1 ? stored.sortKey : 'name',
      hideDone: !!stored.hideDone,
      search: typeof stored.search === 'string' ? stored.search : '',
    };
  });
  const { sortKey, hideDone, search } = prefs;
  const updatePrefs = (patch: Partial<IPortfolioPrefs>): void => {
    setPrefs(prev => {
      const next = { ...prev, ...patch };
      writeStored(PREFS_KEY, next);
      return next;
    });
  };
  const setSortKey = (k: SortKey): void => updatePrefs({ sortKey: k });

  const filteredProjects = React.useMemo(() => {
    const q = search.trim().toLowerCase();
    return projects.filter(p => {
      if (hideDone && (p.status === 'Completed' || p.status === 'Cancelled')) return false;
      if (!q) return true;
      return p.title.toLowerCase().indexOf(q) !== -1
        || (p.description || '').toLowerCase().indexOf(q) !== -1
        || (p.projectManager || '').toLowerCase().indexOf(q) !== -1
        || getProjectStatusLabel(p.status).toLowerCase().indexOf(q) !== -1;
    });
  }, [projects, search, hideDone]);
  const filtersActive = hideDone || !!search.trim();

  const sortedProjects = React.useMemo(() => {
    const list = [...filteredProjects];
    switch (sortKey) {
      case 'name':
        return list.sort((a, b) => a.title.localeCompare(b.title));
      case 'health': {
        return list.sort((a, b) => {
          const ha = statsMap?.get(a.id)?.health ?? 'on-track';
          const hb = statsMap?.get(b.id)?.health ?? 'on-track';
          return HEALTH_ORDER[ha] - HEALTH_ORDER[hb];
        });
      }
      case 'status':
        // Lifecycle order (Planning to Cancelled), not alphabetical.
        return list.sort((a, b) =>
          PROJECT_STATUS_OPTIONS.indexOf(a.status) - PROJECT_STATUS_OPTIONS.indexOf(b.status)
          || a.title.localeCompare(b.title));
      case 'completion': {
        return list.sort((a, b) => {
          const pa = statsMap?.get(a.id)?.overallPct ?? 0;
          const pb = statsMap?.get(b.id)?.overallPct ?? 0;
          return pb - pa;
        });
      }
      default:
        return list;
    }
  }, [filteredProjects, sortKey, statsMap]);

  // Aggregate health summary across all loaded stats
  const summary = React.useMemo(() => {
    if (!statsMap) return null;
    let onTrack = 0, atRisk = 0, overdue = 0, complete = 0;
    statsMap.forEach(s => {
      if (s.statsError) return;
      if (s.health === 'on-track') onTrack++;
      else if (s.health === 'at-risk') atRisk++;
      else if (s.health === 'overdue') overdue++;
      else if (s.health === 'complete') complete++;
    });
    return { onTrack, atRisk, overdue, complete };
  }, [statsMap]);

  const sortOptions: { id: SortKey; label: string }[] = [
    { id: 'name', label: strings.Portfolio_SortAZ },
    { id: 'health', label: strings.Portfolio_SortHealth },
    { id: 'status', label: strings.Portfolio_SortStatus },
    { id: 'completion', label: strings.Portfolio_SortCompletion },
  ];

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden' }}>
      {/* Header bar */}
      <div style={{
        display: 'flex',
        alignItems: 'center',
        gap: 16,
        padding: '12px 20px',
        borderBottom: '1px solid var(--neutralLight, #EDEBE9)',
        background: 'var(--neutralLighterAlt, #FAF9F8)',
        flexShrink: 0,
        flexWrap: 'wrap',
      }}>
        <div style={{ fontWeight: 600, fontSize: 14, color: 'var(--neutralPrimary, #1B1B1B)' }}>
          {filtersActive
            ? formatString(strings.View_Portfolio_FilteredCount, { shown: filteredProjects.length, total: projects.length })
            : formatString(projects.length === 1 ? strings.View_Portfolio_ProjectCountOne : strings.View_Portfolio_ProjectCountMany, { count: projects.length })}
        </div>

        {summary && (
          <div style={{ display: 'flex', gap: 12, fontSize: 12 }}>
            {summary.onTrack > 0 && (
              <span style={{ color: 'var(--themePrimary, #0078D4)' }}>● {formatString(strings.Portfolio_OnTrackSummary, { count: summary.onTrack })}</span>
            )}
            {summary.atRisk > 0 && (
              <span style={{ color: '#CA5010' }}>● {formatString(strings.Portfolio_AtRiskSummary, { count: summary.atRisk })}</span>
            )}
            {summary.overdue > 0 && (
              <span style={{ color: '#D13438' }}>● {formatString(strings.Portfolio_OverdueSummary, { count: summary.overdue })}</span>
            )}
            {summary.complete > 0 && (
              <span style={{ color: '#107C10' }}>● {formatString(strings.Portfolio_DoneSummary, { count: summary.complete })}</span>
            )}
          </div>
        )}

        <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
          <input
            type="search"
            value={search}
            onChange={e => updatePrefs({ search: e.target.value })}
            placeholder={strings.View_Portfolio_SearchPlaceholder}
            aria-label={strings.View_Portfolio_SearchAriaLabel}
            style={{
              padding: '4px 8px', borderRadius: 4, border: '1px solid var(--neutralQuaternary, #C8C6C4)', fontSize: 12,
              width: 170, background: 'var(--white, #fff)', color: 'var(--neutralPrimary, #323130)',
            }}
          />
          <label style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 12, color: 'var(--neutralPrimary, #323130)', cursor: 'pointer' }}>
            <input
              type="checkbox"
              checked={hideDone}
              onChange={e => updatePrefs({ hideDone: e.target.checked })}
            />
            {strings.View_Portfolio_HideDone}
          </label>
          <span style={{ fontSize: 12, color: 'var(--neutralSecondary, #605E5C)' }}>{strings.Portfolio_SortLabel}</span>
          {sortOptions.map(o => (
            <button
              key={o.id}
              onClick={() => setSortKey(o.id)}
              aria-pressed={sortKey === o.id}
              style={{
                padding: '4px 10px',
                borderRadius: 4,
                border: '1px solid',
                borderColor: sortKey === o.id ? 'var(--themePrimary, #0078D4)' : 'var(--neutralLight, #EDEBE9)',
                background: sortKey === o.id ? 'var(--themeLighterAlt, #EFF6FC)' : 'var(--white, #fff)',
                color: sortKey === o.id ? 'var(--themePrimary, #0078D4)' : 'var(--neutralPrimary, #323130)',
                fontSize: 12,
                fontWeight: sortKey === o.id ? 600 : 400,
                cursor: 'pointer',
              }}
            >
              {o.label}
            </button>
          ))}

          <button
            onClick={onRefresh}
            disabled={loading}
            title={strings.Portfolio_RefreshTitle}
            style={{
              padding: '4px 10px',
              borderRadius: 4,
              border: '1px solid var(--neutralLight, #EDEBE9)',
              background: 'var(--white, #fff)',
              fontSize: 12,
              cursor: loading ? 'default' : 'pointer',
              color: 'var(--neutralSecondary, #605E5C)',
            }}
          >
            {loading ? '⟳' : '↻'} {strings.Portfolio_RefreshButton}
          </button>
        </div>
      </div>

      {/* Cards grid */}
      <div style={{ flex: 1, overflowY: 'auto', padding: 20 }}>
        {loading && !statsMap && (
          <div style={{ display: 'flex', justifyContent: 'center', padding: 40 }}>
            <Spinner size={SpinnerSize.large} label={strings.Portfolio_LoadingStats} />
          </div>
        )}

        {!loading && projects.length === 0 && (
          <div style={{ textAlign: 'center', padding: 60 }}>
            <div style={{ fontSize: 48, opacity: 0.3, marginBottom: 12 }}>📋</div>
            <div style={{ fontSize: 16, fontWeight: 600, color: 'var(--neutralPrimary, #323130)', marginBottom: 8 }}>{strings.Portfolio_EmptyTitle}</div>
            <button
              onClick={onAddProject}
              style={{
                background: 'var(--themePrimary, #0078D4)', color: '#fff', border: 'none', borderRadius: 4,
                padding: '8px 16px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
              }}
            >
              {strings.Portfolio_CreateFirstProject}
            </button>
          </div>
        )}

        {projects.length > 0 && filteredProjects.length === 0 && (
          <div style={{ textAlign: 'center', padding: 40, color: 'var(--neutralSecondary, #605E5C)', fontSize: 13 }}>
            <div style={{ marginBottom: 12 }}>{strings.View_Portfolio_NoMatches}</div>
            <button
              onClick={() => updatePrefs({ search: '', hideDone: false })}
              style={{
                padding: '4px 12px', borderRadius: 4, border: '1px solid var(--neutralQuaternary, #C8C6C4)',
                background: 'var(--white, #fff)', color: 'var(--neutralPrimary, #323130)', fontSize: 12, cursor: 'pointer',
              }}
            >
              {strings.View_Portfolio_ClearFilters}
            </button>
          </div>
        )}

        {filteredProjects.length > 0 && (
          <div style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))',
            gap: 16,
          }}>
            {sortedProjects.map(project => (
              <ProjectCard
                key={project.id}
                project={project}
                stats={statsMap?.get(project.id)}
                statsLoading={loading}
                onClick={() => onSelectProject(project)}
              />
            ))}
          </div>
        )}
      </div>

      <style>{`
        @keyframes pulse {
          0%, 100% { opacity: 1; }
          50% { opacity: 0.5; }
        }
        ${SHARED_VIEW_CSS}
      `}</style>
    </div>
  );
};
