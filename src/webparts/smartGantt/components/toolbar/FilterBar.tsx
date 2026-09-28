import * as React from 'react';
import { Callout, Checkbox, DirectionalHint } from '@fluentui/react';
import {
  ITaskFilter, isFilterActive, DueFilter,
  TASK_STATUS_OPTIONS, TASK_PRIORITY_OPTIONS, TaskStatus, TaskPriority,
} from '../../models';
import * as strings from 'SmartGanttWebPartStrings';
import { formatString } from '../localeUtils';

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
        {active
          ? formatString(strings.FilterBar_ChipLabelWithCount, { label, count: selected.length })
          : formatString(strings.FilterBar_ChipLabelNoCount, { label })}
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
                {strings.FilterBar_ClearOption}
              </button>
            )}
          </div>
        </Callout>
      )}
    </>
  );
}

const DUE_OPTIONS: { id: DueFilter; label: string }[] = [
  { id: 'all', label: strings.FilterBar_DueAny },
  { id: 'overdue', label: strings.FilterBar_DueOverdue },
  { id: 'today', label: strings.FilterBar_DueToday },
  { id: 'week', label: strings.FilterBar_DueWeek },
];

const FilterBarComponent: React.FC<IFilterBarProps> = ({
  filter, onChange, assignees, phases, matchCount, totalCount,
}) => {
  const active = isFilterActive(filter);
  const set = <K extends keyof ITaskFilter>(key: K, value: ITaskFilter[K]): void => {
    onChange({ ...filter, [key]: value });
  };

  // The search box has local state debounced ~150ms before it reaches the
  // parent — without it, every keystroke re-renders the toolbar plus the
  // full Gantt/List/Kanban view underneath.
  const [localText, setLocalText] = React.useState(filter.text);
  const debounceRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  // Keep local text in sync when the filter changes externally (project
  // switch, "Clear filters").
  React.useEffect(() => setLocalText(filter.text), [filter.text]);

  React.useEffect(() => () => {
    if (debounceRef.current) clearTimeout(debounceRef.current);
  }, []);

  const handleTextChange = (v: string): void => {
    setLocalText(v);
    if (debounceRef.current) clearTimeout(debounceRef.current);
    debounceRef.current = setTimeout(() => set('text', v), 150);
  };

  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
      <input
        type="search"
        value={localText}
        onChange={e => handleTextChange(e.target.value)}
        placeholder={strings.FilterBar_SearchPlaceholder}
        aria-label={strings.FilterBar_SearchAriaLabel}
        style={{
          width: 160,
          padding: '4px 10px',
          borderRadius: 12,
          border: `1px solid ${localText ? '#0078D4' : '#D2D0CE'}`,
          fontSize: 12,
          outline: 'none',
          fontFamily: 'inherit',
        }}
      />
      <MultiChip<TaskStatus>
        label={strings.FilterBar_StatusChipLabel}
        options={TASK_STATUS_OPTIONS}
        selected={filter.statuses}
        onChange={v => set('statuses', v)}
      />
      <MultiChip<TaskPriority>
        label={strings.FilterBar_PriorityChipLabel}
        options={TASK_PRIORITY_OPTIONS}
        selected={filter.priorities}
        onChange={v => set('priorities', v)}
      />
      {assignees.length > 0 && (
        <MultiChip<string>
          label={strings.FilterBar_AssigneeChipLabel}
          options={assignees}
          selected={filter.assignees}
          onChange={v => set('assignees', v)}
        />
      )}
      {phases.length > 0 && (
        <MultiChip<string>
          label={strings.FilterBar_PhaseChipLabel}
          options={phases}
          selected={filter.phases}
          onChange={v => set('phases', v)}
        />
      )}
      <select
        value={filter.due}
        onChange={e => set('due', e.target.value as DueFilter)}
        aria-label={strings.FilterBar_DueDateAriaLabel}
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
            {formatString(strings.FilterBar_MatchCount, { matchCount, totalCount })}
          </span>
          <button
            style={{
              background: 'none', border: 'none', color: '#0078D4',
              fontSize: 12, fontWeight: 600, cursor: 'pointer', padding: '3px 6px',
            }}
            onClick={() => {
              if (debounceRef.current) clearTimeout(debounceRef.current);
              setLocalText('');
              onChange({ text: '', statuses: [], priorities: [], assignees: [], phases: [], due: 'all' });
            }}
          >
            {strings.FilterBar_ClearFilters}
          </button>
        </>
      )}
    </div>
  );
};

export const FilterBar = React.memo(FilterBarComponent);

export default FilterBar;
