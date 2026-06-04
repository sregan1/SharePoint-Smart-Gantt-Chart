import * as React from 'react';
import { format, addDays, startOfDay, differenceInCalendarDays } from 'date-fns';
import {
  IProject, ITask, TaskStatus,
  STATUS_COLORS, STATUS_LIGHT_COLORS, PRIORITY_COLORS, phaseColor,
} from '../../models';

interface IDashboardViewProps {
  project: IProject;
  tasks: ITask[];
  onEditTask: (task: ITask) => void;
  onAddTask: () => void;
}

function parseDate(s: string): Date | null {
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : startOfDay(d);
}

function fmtDate(s: string): string {
  const d = parseDate(s);
  return d ? format(d, 'MMM d') : '—';
}

// ── Stat card ──────────────────────────────────────────────────────────────
const StatCard: React.FC<{
  label: string; value: number; total: number;
  color: string; bg: string; onClick?: () => void;
}> = ({ label, value, total, color, bg, onClick }) => (
  <div
    onClick={onClick}
    style={{
      flex: 1, minWidth: 110,
      background: bg, border: `1px solid ${color}30`,
      borderRadius: 8, padding: '14px 16px',
      cursor: onClick ? 'pointer' : 'default',
      transition: 'box-shadow 0.15s',
    }}
    onMouseEnter={e => { if (onClick) (e.currentTarget as HTMLElement).style.boxShadow = `0 2px 8px ${color}30`; }}
    onMouseLeave={e => { (e.currentTarget as HTMLElement).style.boxShadow = 'none'; }}
  >
    <div style={{ fontSize: 30, fontWeight: 700, color, lineHeight: 1 }}>{value}</div>
    <div style={{ fontSize: 11, color: '#605E5C', marginTop: 4 }}>{label}</div>
    {total > 0 && (
      <div style={{ marginTop: 8, height: 4, background: `${color}25`, borderRadius: 2, overflow: 'hidden' }}>
        <div style={{ height: '100%', width: `${Math.round(value / total * 100)}%`, background: color, borderRadius: 2 }} />
      </div>
    )}
  </div>
);

// ── Section header ─────────────────────────────────────────────────────────
const SectionHeader: React.FC<{ title: string }> = ({ title }) => (
  <div style={{
    fontSize: 11, fontWeight: 700, color: '#605E5C',
    letterSpacing: '0.6px', textTransform: 'uppercase',
    paddingBottom: 8, borderBottom: '1px solid #EDEBE9', marginBottom: 14,
  }}>
    {title}
  </div>
);

// ── Task row ───────────────────────────────────────────────────────────────
const TaskRow: React.FC<{ task: ITask; onClick: () => void; showDue?: boolean; isOverdue?: boolean }> = ({
  task, onClick, showDue, isOverdue,
}) => {
  const sc = STATUS_COLORS[task.status as TaskStatus] || '#8B929A';
  return (
    <div
      onClick={onClick}
      style={{
        display: 'flex', alignItems: 'center', gap: 10,
        padding: '7px 10px', borderRadius: 5, cursor: 'pointer',
        background: isOverdue ? '#FDF3F4' : 'transparent',
        border: isOverdue ? '1px solid #F9D8D8' : '1px solid transparent',
        marginBottom: 4, transition: 'background 0.1s',
      }}
      onMouseEnter={e => { (e.currentTarget as HTMLElement).style.background = isOverdue ? '#FAE5E5' : '#F3F2F1'; }}
      onMouseLeave={e => { (e.currentTarget as HTMLElement).style.background = isOverdue ? '#FDF3F4' : 'transparent'; }}
    >
      <div style={{ width: 8, height: 8, borderRadius: '50%', background: sc, flexShrink: 0 }} />
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontSize: 13, color: '#323130', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
          {task.isMilestone ? '◆ ' : ''}{task.title}
        </div>
        <div style={{ fontSize: 11, color: '#605E5C', marginTop: 1 }}>
          <span style={{ color: sc }}>{task.status}</span>
          {task.percentComplete > 0 && <span>  •  {task.percentComplete}%</span>}
          {task.phase && <span>  •  {task.phase}</span>}
        </div>
      </div>
      {showDue && (
        <div style={{ fontSize: 11, fontWeight: 600, color: isOverdue ? '#D13438' : '#605E5C', flexShrink: 0 }}>
          {isOverdue ? '⚠ ' : ''}{fmtDate(task.dueDate)}
        </div>
      )}
      {task.percentComplete > 0 && !showDue && (
        <div style={{ width: 60, flexShrink: 0 }}>
          <div style={{ height: 4, background: '#EDEBE9', borderRadius: 2, overflow: 'hidden' }}>
            <div style={{ height: '100%', width: `${task.percentComplete}%`, background: sc, borderRadius: 2 }} />
          </div>
        </div>
      )}
    </div>
  );
};

// ── Main component ─────────────────────────────────────────────────────────
export const DashboardView: React.FC<IDashboardViewProps> = ({
  project, tasks, onEditTask, onAddTask,
}) => {
  const today = startOfDay(new Date());
  const in14 = addDays(today, 14);
  const weekAgo = addDays(today, -7);

  // ── Aggregate stats ──────────────────────────────────────────────────────
  const total = tasks.length;
  const byStatus: Record<string, number> = {
    'Completed': 0, 'In Progress': 0, 'Not Started': 0, 'On Hold': 0, 'Cancelled': 0,
  };
  tasks.forEach(t => { if (byStatus[t.status] !== undefined) byStatus[t.status]++; });
  const overallPct = total > 0
    ? Math.round(tasks.reduce((s, t) => s + t.percentComplete, 0) / total)
    : 0;

  // ── Phase progress ───────────────────────────────────────────────────────
  const phaseMap = new Map<string, ITask[]>();
  tasks.forEach(t => {
    if (!t.phase) return;
    if (!phaseMap.has(t.phase)) phaseMap.set(t.phase, []);
    phaseMap.get(t.phase)!.push(t);
  });
  const phases = Array.from(phaseMap.entries()).map(([name, pts]) => ({
    name,
    color: phaseColor(name),
    total: pts.length,
    completed: pts.filter(t => t.status === 'Completed').length,
    pct: Math.round(pts.reduce((s, t) => s + t.percentComplete, 0) / pts.length),
  }));

  // ── Time-based groups ────────────────────────────────────────────────────
  const recentlyModified = tasks.filter(t => {
    const mod = t.modified ? parseDate(t.modified) : null;
    return mod !== null && mod >= weekAgo && t.status !== 'Completed';
  }).slice(0, 6);

  const recentlyCompleted = tasks.filter(t => {
    const mod = t.modified ? parseDate(t.modified) : null;
    return mod !== null && mod >= weekAgo && t.status === 'Completed';
  }).slice(0, 4);

  const upcoming = tasks.filter(t => {
    const due = parseDate(t.dueDate);
    return due !== null && due >= today && due <= in14 && t.status !== 'Completed' && t.status !== 'Cancelled';
  }).sort((a, b) => (a.dueDate > b.dueDate ? 1 : -1)).slice(0, 6);

  const overdue = tasks.filter(t => {
    const due = parseDate(t.dueDate);
    return due !== null && due < today && t.status !== 'Completed' && t.status !== 'Cancelled';
  }).sort((a, b) => (a.dueDate < b.dueDate ? 1 : -1)).slice(0, 6);

  // ── Priority breakdown ───────────────────────────────────────────────────
  const byPriority: Record<string, number> = { Critical: 0, High: 0, Medium: 0, Low: 0 };
  tasks.forEach(t => { if (byPriority[t.priority] !== undefined) byPriority[t.priority]++; });

  const priorityItems = [
    { label: 'Critical', color: PRIORITY_COLORS['Critical'] },
    { label: 'High',     color: PRIORITY_COLORS['High']     },
    { label: 'Medium',   color: PRIORITY_COLORS['Medium']   },
    { label: 'Low',      color: PRIORITY_COLORS['Low']      },
  ];

  return (
    <div style={{ padding: '20px 24px', overflowY: 'auto', background: '#FAF9F8', minHeight: '100%' }}>

      {/* ── Project header ─────────────────────────────────────────────── */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 20 }}>
        <div style={{ width: 14, height: 14, borderRadius: '50%', background: project.color, flexShrink: 0 }} />
        <div>
          <div style={{ fontSize: 20, fontWeight: 700, color: '#323130', lineHeight: 1.2 }}>{project.title}</div>
          {project.description && (
            <div style={{ fontSize: 12, color: '#605E5C', marginTop: 2 }}>{project.description}</div>
          )}
        </div>
        <div style={{ marginLeft: 'auto', display: 'flex', gap: 8 }}>
          {project.startDate && project.dueDate && (
            <div style={{ fontSize: 12, color: '#605E5C', background: '#fff', border: '1px solid #EDEBE9', borderRadius: 4, padding: '4px 10px' }}>
              {fmtDate(project.startDate)} → {fmtDate(project.dueDate)}
            </div>
          )}
          <div style={{
            fontSize: 12, fontWeight: 600, padding: '4px 10px', borderRadius: 4,
            background: STATUS_LIGHT_COLORS[project.status as TaskStatus] || '#F3F2F1',
            color: STATUS_COLORS[project.status as TaskStatus] || '#605E5C',
          }}>
            {project.status}
          </div>
        </div>
      </div>

      {/* ── Stat cards ─────────────────────────────────────────────────── */}
      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginBottom: 20 }}>
        <StatCard label="Total Tasks"  value={total}                   total={0}     color="#323130" bg="#F3F2F1" />
        <StatCard label="Completed"    value={byStatus['Completed']}   total={total} color={STATUS_COLORS['Completed']}   bg={STATUS_LIGHT_COLORS['Completed']}   />
        <StatCard label="In Progress"  value={byStatus['In Progress']} total={total} color={STATUS_COLORS['In Progress']} bg={STATUS_LIGHT_COLORS['In Progress']} />
        <StatCard label="On Hold"      value={byStatus['On Hold']}     total={total} color={STATUS_COLORS['On Hold']}     bg={STATUS_LIGHT_COLORS['On Hold']}     />
        <StatCard label="Not Started"  value={byStatus['Not Started']} total={total} color={STATUS_COLORS['Not Started']} bg={STATUS_LIGHT_COLORS['Not Started']} />
        {byStatus['Cancelled'] > 0 && (
          <StatCard label="Cancelled"  value={byStatus['Cancelled']}   total={total} color={STATUS_COLORS['Cancelled']}   bg={STATUS_LIGHT_COLORS['Cancelled']}   />
        )}
      </div>

      {/* ── Overall progress ───────────────────────────────────────────── */}
      <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #EDEBE9', padding: '16px 20px', marginBottom: 16 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
          <span style={{ fontSize: 13, fontWeight: 700, color: '#323130' }}>Overall Progress</span>
          <span style={{ fontSize: 16, fontWeight: 700, color: project.color }}>{overallPct}%</span>
        </div>
        <div style={{ height: 10, background: '#EDEBE9', borderRadius: 5, overflow: 'hidden' }}>
          <div style={{ height: '100%', width: `${overallPct}%`, background: project.color, borderRadius: 5, transition: 'width 0.4s ease' }} />
        </div>
        {project.projectManager && (
          <div style={{ fontSize: 11, color: '#8A8886', marginTop: 8 }}>Project Manager: {project.projectManager}</div>
        )}
      </div>

      {/* ── Phase progress + Status / Priority breakdown ────────────────── */}
      <div style={{ display: 'grid', gridTemplateColumns: phases.length ? '1fr 1fr' : '1fr', gap: 16, marginBottom: 16 }}>

        {/* Phase progress */}
        {phases.length > 0 && (
          <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #EDEBE9', padding: '16px 20px' }}>
            <SectionHeader title="Phase Progress" />
            {phases.map(ph => (
              <div key={ph.name} style={{ marginBottom: 12 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 4 }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <div style={{ width: 8, height: 8, borderRadius: '50%', background: ph.color }} />
                    <span style={{ fontSize: 12, color: '#323130', fontWeight: 500 }}>{ph.name}</span>
                  </div>
                  <span style={{ fontSize: 11, color: '#605E5C' }}>{ph.completed}/{ph.total} done  •  {ph.pct}%</span>
                </div>
                <div style={{ height: 6, background: '#EDEBE9', borderRadius: 3, overflow: 'hidden' }}>
                  <div style={{ height: '100%', width: `${ph.pct}%`, background: ph.color, borderRadius: 3 }} />
                </div>
              </div>
            ))}
          </div>
        )}

        {/* Status + Priority breakdown */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
          <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #EDEBE9', padding: '16px 20px', flex: 1 }}>
            <SectionHeader title="Status Breakdown" />
            {(['Completed','In Progress','Not Started','On Hold','Cancelled'] as TaskStatus[]).map(s => {
              const cnt = byStatus[s] || 0;
              const pct = total > 0 ? Math.round(cnt / total * 100) : 0;
              return (
                <div key={s} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 7 }}>
                  <div style={{ width: 8, height: 8, borderRadius: '50%', background: STATUS_COLORS[s], flexShrink: 0 }} />
                  <span style={{ fontSize: 12, color: '#323130', flex: 1 }}>{s}</span>
                  <span style={{ fontSize: 12, fontWeight: 600, color: STATUS_COLORS[s], width: 24, textAlign: 'right' }}>{cnt}</span>
                  <div style={{ width: 80, height: 4, background: '#EDEBE9', borderRadius: 2, overflow: 'hidden' }}>
                    <div style={{ height: '100%', width: `${pct}%`, background: STATUS_COLORS[s], borderRadius: 2 }} />
                  </div>
                  <span style={{ fontSize: 10, color: '#8A8886', width: 30, textAlign: 'right' }}>{pct}%</span>
                </div>
              );
            })}
          </div>

          <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #EDEBE9', padding: '16px 20px' }}>
            <SectionHeader title="Priority Breakdown" />
            {priorityItems.map(p => {
              const cnt = byPriority[p.label] || 0;
              const pct = total > 0 ? Math.round(cnt / total * 100) : 0;
              return (
                <div key={p.label} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 7 }}>
                  <div style={{ width: 8, height: 8, borderRadius: '50%', background: p.color, flexShrink: 0 }} />
                  <span style={{ fontSize: 12, color: '#323130', flex: 1 }}>{p.label}</span>
                  <span style={{ fontSize: 12, fontWeight: 600, color: p.color, width: 24, textAlign: 'right' }}>{cnt}</span>
                  <div style={{ width: 80, height: 4, background: '#EDEBE9', borderRadius: 2, overflow: 'hidden' }}>
                    <div style={{ height: '100%', width: `${pct}%`, background: p.color, borderRadius: 2 }} />
                  </div>
                  <span style={{ fontSize: 10, color: '#8A8886', width: 30, textAlign: 'right' }}>{pct}%</span>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* ── Overdue + Upcoming ──────────────────────────────────────────── */}
      {(overdue.length > 0 || upcoming.length > 0) && (
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 16 }}>

          {overdue.length > 0 && (
            <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #EDEBE9', padding: '16px 20px' }}>
              <SectionHeader title={`Overdue  (${overdue.length})`} />
              {overdue.map(t => (
                <TaskRow key={t.id} task={t} onClick={() => onEditTask(t)} showDue isOverdue />
              ))}
            </div>
          )}

          {upcoming.length > 0 && (
            <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #EDEBE9', padding: '16px 20px' }}>
              <SectionHeader title={`Due in next 14 days  (${upcoming.length})`} />
              {upcoming.map(t => {
                const due = parseDate(t.dueDate)!;
                const daysLeft = differenceInCalendarDays(due, today);
                return (
                  <div key={t.id} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4, cursor: 'pointer', padding: '6px 8px', borderRadius: 5 }}
                    onClick={() => onEditTask(t)}
                    onMouseEnter={e => { (e.currentTarget as HTMLElement).style.background = '#F3F2F1'; }}
                    onMouseLeave={e => { (e.currentTarget as HTMLElement).style.background = 'transparent'; }}
                  >
                    <div style={{ width: 8, height: 8, borderRadius: '50%', background: STATUS_COLORS[t.status as TaskStatus] || '#8B929A', flexShrink: 0 }} />
                    <span style={{ fontSize: 13, color: '#323130', flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {t.isMilestone ? '◆ ' : ''}{t.title}
                    </span>
                    <span style={{ fontSize: 11, color: daysLeft <= 3 ? '#CA5010' : '#605E5C', fontWeight: daysLeft <= 3 ? 600 : 400, flexShrink: 0 }}>
                      {daysLeft === 0 ? 'Today' : daysLeft === 1 ? 'Tomorrow' : `${daysLeft}d`}
                    </span>
                  </div>
                );
              })}
            </div>
          )}

        </div>
      )}

      {/* ── Recent activity ─────────────────────────────────────────────── */}
      {(recentlyCompleted.length > 0 || recentlyModified.length > 0) && (
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 16 }}>

          {recentlyCompleted.length > 0 && (
            <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #EDEBE9', padding: '16px 20px' }}>
              <SectionHeader title="Completed this week" />
              {recentlyCompleted.map(t => (
                <TaskRow key={t.id} task={t} onClick={() => onEditTask(t)} />
              ))}
            </div>
          )}

          {recentlyModified.length > 0 && (
            <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #EDEBE9', padding: '16px 20px' }}>
              <SectionHeader title="Updated this week" />
              {recentlyModified.map(t => (
                <TaskRow key={t.id} task={t} onClick={() => onEditTask(t)} />
              ))}
            </div>
          )}

        </div>
      )}

      {/* ── Empty state ─────────────────────────────────────────────────── */}
      {total === 0 && (
        <div style={{ textAlign: 'center', padding: '60px 0', color: '#605E5C' }}>
          <div style={{ fontSize: 32, marginBottom: 12 }}>📋</div>
          <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 8, color: '#323130' }}>No tasks yet</div>
          <div style={{ fontSize: 13, marginBottom: 20 }}>Add your first task to see the project summary.</div>
          <button
            onClick={onAddTask}
            style={{
              padding: '8px 20px', background: project.color, color: '#fff',
              border: 'none', borderRadius: 4, fontSize: 13, fontWeight: 600, cursor: 'pointer',
            }}
          >
            + Add Task
          </button>
        </div>
      )}

    </div>
  );
};

export default DashboardView;
