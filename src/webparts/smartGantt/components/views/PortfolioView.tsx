import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { differenceInCalendarDays } from 'date-fns';
import { IProject, IProjectTaskStats } from '../../models';
import { HealthBadge } from '../common/HealthBadge';
import { ProjectHealth } from '../../models';
import { parseDateOnly, formatDateOnly, todayLocalMidnight } from '../../utils/dateUtils';

// ─── Types ────────────────────────────────────────────────────────────────────

export interface IPortfolioViewProps {
  projects: IProject[];
  statsMap: Map<number, IProjectTaskStats> | null;
  loading: boolean;
  onSelectProject: (project: IProject) => void;
  onAddProject: () => void;
  onRefresh: () => void;
}

type SortKey = 'name' | 'health' | 'status' | 'completion';

const HEALTH_ORDER: Record<ProjectHealth, number> = { overdue: 0, 'at-risk': 1, 'on-track': 2, complete: 3 };

const STATUS_COLORS: Record<string, string> = {
  Planning: '#8764B8',
  Active: '#0078D4',
  'On Hold': '#CA5010',
  Completed: '#107C10',
  Cancelled: '#8B929A',
};

const STATUS_LIGHT_COLORS: Record<string, string> = {
  Planning: '#F3EFF8',
  Active: '#EFF6FC',
  'On Hold': '#FFF4EC',
  Completed: '#F1FAF1',
  Cancelled: '#F3F2F1',
};

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

  const total = differenceInCalendarDays(endDate, startDate);
  if (total <= 0) return null;

  const elapsed = Math.max(0, Math.min(total, differenceInCalendarDays(today, startDate)));
  const todayPct = Math.round((elapsed / total) * 100);
  const clampedPct = Math.max(0, Math.min(100, todayPct));

  return (
    <div style={{ marginTop: 12 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 11, color: '#605E5C', marginBottom: 4 }}>
        <span>{formatDate(start)}</span>
        <span>{formatDate(end)}</span>
      </div>
      <div style={{ position: 'relative', height: 6, background: '#EDEBE9', borderRadius: 3, overflow: 'visible' }}>
        <div style={{
          position: 'absolute',
          left: 0,
          top: 0,
          height: '100%',
          width: `${clampedPct}%`,
          background: `${color}40`,
          borderRadius: 3,
        }} />
        {clampedPct > 0 && clampedPct < 100 && (
          <div style={{
            position: 'absolute',
            left: `${clampedPct}%`,
            top: -4,
            width: 2,
            height: 14,
            background: '#D13438',
            borderRadius: 1,
            transform: 'translateX(-50%)',
          }}
            title="Today"
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
  const statusColor = STATUS_COLORS[project.status] || '#8B929A';
  const statusBg = STATUS_LIGHT_COLORS[project.status] || '#F3F2F1';

  return (
    <div
      onClick={onClick}
      style={{
        background: '#fff',
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
          <div style={{ fontSize: 15, fontWeight: 600, color: '#1B1B1B', marginBottom: 4, lineHeight: 1.3 }}>
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
              {project.status}
            </span>
            {stats && <HealthBadge health={stats.health} size="md" />}
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
        <div style={{ height: 32, background: '#F3F2F1', borderRadius: 4, animation: 'pulse 1.5s ease-in-out infinite' }} />
      ) : stats ? (
        <>
          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, color: '#605E5C', marginBottom: 4 }}>
            <span>Overall Progress</span>
            <span style={{ fontWeight: 600, color: '#1B1B1B' }}>{stats.overallPct}%</span>
          </div>
          <div style={{ height: 8, background: '#EDEBE9', borderRadius: 4, overflow: 'hidden', marginBottom: 12 }}>
            <div style={{
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
              { label: 'Done', value: stats.completedCount, color: '#107C10' },
              { label: 'Active', value: stats.inProgressCount, color: '#0078D4' },
              { label: 'At Risk', value: stats.atRiskCount, color: '#CA5010' },
              { label: 'Overdue', value: stats.overdueCount, color: '#D13438' },
            ] as const).map(({ label, value, color }) => (
              <div key={label} style={{ textAlign: 'center', padding: '6px 4px', background: '#FAF9F8', borderRadius: 6 }}>
                <div style={{ fontSize: 18, fontWeight: 700, color, lineHeight: 1 }}>{value}</div>
                <div style={{ fontSize: 10, color: '#605E5C', marginTop: 2 }}>{label}</div>
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
      <div style={{ marginTop: 12, paddingTop: 10, borderTop: '1px solid #F3F2F1', display: 'flex', justifyContent: 'flex-end' }}>
        <span style={{ fontSize: 12, color: '#0078D4', fontWeight: 600 }}>
          View Gantt →
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
  const [sortKey, setSortKey] = React.useState<SortKey>('name');

  const sortedProjects = React.useMemo(() => {
    const list = [...projects];
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
        return list.sort((a, b) => a.status.localeCompare(b.status));
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
  }, [projects, sortKey, statsMap]);

  // Aggregate health summary across all loaded stats
  const summary = React.useMemo(() => {
    if (!statsMap) return null;
    let onTrack = 0, atRisk = 0, overdue = 0, complete = 0;
    statsMap.forEach(s => {
      if (s.health === 'on-track') onTrack++;
      else if (s.health === 'at-risk') atRisk++;
      else if (s.health === 'overdue') overdue++;
      else if (s.health === 'complete') complete++;
    });
    return { onTrack, atRisk, overdue, complete };
  }, [statsMap]);

  const sortOptions: { id: SortKey; label: string }[] = [
    { id: 'name', label: 'A–Z' },
    { id: 'health', label: 'Health' },
    { id: 'status', label: 'Status' },
    { id: 'completion', label: '% Done' },
  ];

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden' }}>
      {/* Header bar */}
      <div style={{
        display: 'flex',
        alignItems: 'center',
        gap: 16,
        padding: '12px 20px',
        borderBottom: '1px solid #EDEBE9',
        background: '#FAF9F8',
        flexShrink: 0,
        flexWrap: 'wrap',
      }}>
        <div style={{ fontWeight: 600, fontSize: 14, color: '#1B1B1B' }}>
          {projects.length} Project{projects.length !== 1 ? 's' : ''}
        </div>

        {summary && (
          <div style={{ display: 'flex', gap: 12, fontSize: 12 }}>
            {summary.onTrack > 0 && (
              <span style={{ color: '#0078D4' }}>● {summary.onTrack} On Track</span>
            )}
            {summary.atRisk > 0 && (
              <span style={{ color: '#CA5010' }}>● {summary.atRisk} At Risk</span>
            )}
            {summary.overdue > 0 && (
              <span style={{ color: '#D13438' }}>● {summary.overdue} Overdue</span>
            )}
            {summary.complete > 0 && (
              <span style={{ color: '#107C10' }}>● {summary.complete} Done</span>
            )}
          </div>
        )}

        <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 8 }}>
          <span style={{ fontSize: 12, color: '#605E5C' }}>Sort:</span>
          {sortOptions.map(o => (
            <button
              key={o.id}
              onClick={() => setSortKey(o.id)}
              style={{
                padding: '4px 10px',
                borderRadius: 4,
                border: '1px solid',
                borderColor: sortKey === o.id ? '#0078D4' : '#EDEBE9',
                background: sortKey === o.id ? '#EFF6FC' : '#fff',
                color: sortKey === o.id ? '#0078D4' : '#323130',
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
            title="Refresh stats"
            style={{
              padding: '4px 10px',
              borderRadius: 4,
              border: '1px solid #EDEBE9',
              background: '#fff',
              fontSize: 12,
              cursor: loading ? 'default' : 'pointer',
              color: '#605E5C',
            }}
          >
            {loading ? '⟳' : '↻'} Refresh
          </button>
        </div>
      </div>

      {/* Cards grid */}
      <div style={{ flex: 1, overflowY: 'auto', padding: 20 }}>
        {loading && !statsMap && (
          <div style={{ display: 'flex', justifyContent: 'center', padding: 40 }}>
            <Spinner size={SpinnerSize.large} label="Loading project stats…" />
          </div>
        )}

        {!loading && projects.length === 0 && (
          <div style={{ textAlign: 'center', padding: 60 }}>
            <div style={{ fontSize: 48, opacity: 0.3, marginBottom: 12 }}>📋</div>
            <div style={{ fontSize: 16, fontWeight: 600, color: '#323130', marginBottom: 8 }}>No projects yet</div>
            <button
              onClick={onAddProject}
              style={{
                background: '#0078D4', color: '#fff', border: 'none', borderRadius: 4,
                padding: '8px 16px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
              }}
            >
              + Create First Project
            </button>
          </div>
        )}

        {projects.length > 0 && (
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
      `}</style>
    </div>
  );
};
