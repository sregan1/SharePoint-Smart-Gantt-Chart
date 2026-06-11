import * as React from 'react';
import { Callout, Checkbox, DirectionalHint } from '@fluentui/react';
import {
  ITaskFilter, isFilterActive, DueFilter,
  TASK_STATUS_OPTIONS, TASK_PRIORITY_OPTIONS, TaskStatus, TaskPriority,
} from '../../models';

interface IFilterBarProps {
  filter: ITaskFilter;
  onChange: (f: ITaskFilter) => void;
  assignees: string[];
  phases: string[];
  matchCount: number;
  totalCount: number;
}

const chipStyle = (active: boolean): React.CSSProperties => ({
  display: 'inline-flex',
  alignItems: 'center',
  gap: 4,
  padding: '3px 10px',
  borderRadius: 12,
  border: `1px solid ${active ? '#0078D4' : '#D2D0CE'}`,
  background: active ? '#EFF6FC' : '#fff',
  color: active ? '#0078D4' : '#605E5C',
  fontSize: 12,
  fontWeight: active ? 600 : 400,
  cursor: 'pointer',
  whiteSpace: 'nowrap',
});

interface IMultiChipProps<T extends string> {
  label: string;
  options: T[];
  selected: T[];
  onChange: (next: T[]) => void;
}

function MultiChip<T extends string>({ label, options, selected, onChange }: IMultiChipProps<T>): React.ReactElement {
  const [open, setOpen] = React.useState(false);
  const ref = React.useRef<HTMLButtonElement>(null);
  const active = selected.length > 0;

  return (
    <>
      <button
        ref={ref}
        style={chipStyle(active)}
        onClick={() => setOpen(v => !v)}
        aria-haspopup="true"
        aria-expanded={open}
      >
        {label}{active ? ` (${selected.length})` : ''} ▾
      </button>
      {open && (
        <Callout
          target={ref}
          onDismiss={() => setOpen(false)}
          directionalHint={DirectionalHint.bottomLeftEdge}
          isBeakVisible={false}
        >
          <div style={{ padding: '10px 14px', display: 'flex', flexDirection: 'column', gap: 8, minWidth: 160 }}>
            {options.map(opt => (
              <Checkbox
                key={opt}
                label={opt}
                checked={selected.indexOf(opt) !== -1}
                onChange={(_, checked) => {
                  onChange(checked ? [...selected, opt] : selected.filter(s => s !== opt));
                }}
              />
            ))}
            {active && (
              <button
                style={{
                  background: 'none', border: 'none', color: '#0078D4', fontSize: 12,
                  cursor: 'pointer', padding: 0, textAlign: 'left',
                }}
                onClick={() => onChange([])}
              >
                Clear
              </button>
            )}
          </div>
        </Callout>
      )}
    </>
  );
}

const DUE_OPTIONS: { id: DueFilter; label: string }[] = [
  { id: 'all', label: 'Any due date' },
  { id: 'overdue', label: 'Overdue' },
  { id: 'today', label: 'Due today' },
  { id: 'week', label: 'Due in 7 days' },
];

export const FilterBar: React.FC<IFilterBarProps> = ({
  filter, onChange, assignees, phases, matchCount, totalCount,
}) => {
  const active = isFilterActive(filter);
  const set = <K extends keyof ITaskFilter>(key: K, value: ITaskFilter[K]): void => {
    onChange({ ...filter, [key]: value });
  };

  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
      <input
        type="search"
        value={filter.text}
        onChange={e => set('text', e.target.value)}
        placeholder="Search tasks…"
        aria-label="Search tasks"
        style={{
          width: 160,
          padding: '4px 10px',
          borderRadius: 12,
          border: `1px solid ${filter.text ? '#0078D4' : '#D2D0CE'}`,
          fontSize: 12,
          outline: 'none',
          fontFamily: 'inherit',
        }}
      />
      <MultiChip<TaskStatus>
        label="Status"
        options={TASK_STATUS_OPTIONS}
        selected={filter.statuses}
        onChange={v => set('statuses', v)}
      />
      <MultiChip<TaskPriority>
        label="Priority"
        options={TASK_PRIORITY_OPTIONS}
        selected={filter.priorities}
        onChange={v => set('priorities', v)}
      />
      {assignees.length > 0 && (
        <MultiChip<string>
          label="Assignee"
          options={assignees}
          selected={filter.assignees}
          onChange={v => set('assignees', v)}
        />
      )}
      {phases.length > 0 && (
        <MultiChip<string>
          label="Phase"
          options={phases}
          selected={filter.phases}
          onChange={v => set('phases', v)}
        />
      )}
      <select
        value={filter.due}
        onChange={e => set('due', e.target.value as DueFilter)}
        aria-label="Filter by due date"
        style={{
          ...chipStyle(filter.due !== 'all'),
          appearance: 'none',
          WebkitAppearance: 'none',
          paddingRight: 18,
        }}
      >
        {DUE_OPTIONS.map(o => (
          <option key={o.id} value={o.id}>{o.label}</option>
        ))}
      </select>
      {active && (
        <>
          <span style={{ fontSize: 12, color: '#605E5C' }}>
            {matchCount} of {totalCount}
          </span>
          <button
            style={{
              background: 'none', border: 'none', color: '#0078D4',
              fontSize: 12, fontWeight: 600, cursor: 'pointer', padding: '3px 6px',
            }}
            onClick={() => onChange({ text: '', statuses: [], priorities: [], assignees: [], phases: [], due: 'all' })}
          >
            ✕ Clear filters
          </button>
        </>
      )}
    </div>
  );
};

export default FilterBar;
