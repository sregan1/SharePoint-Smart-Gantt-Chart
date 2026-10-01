import * as React from 'react';
import {
  Panel, PanelType, TextField, Dropdown, IDropdownOption,
  PrimaryButton, DefaultButton, Stack, Label, Toggle, Spinner, SpinnerSize,
  Slider, MessageBar, MessageBarType,
} from '@fluentui/react';
import {
  ITask, IProject, TaskStatus, TaskPriority, DependencyType, IDependencyLink, IWorkingCalendar,
  DEFAULT_WORKING_CALENDAR,
  TASK_STATUS_OPTIONS, TASK_PRIORITY_OPTIONS,
  STATUS_COLORS, PRIORITY_COLORS, PROJECT_COLORS,
} from '../../models';
import { AutocompleteField } from '../common/AutocompleteField';
import { ColorSwatchPicker } from '../common/ColorSwatchPicker';
import {
  PeoplePicker, IPeopleSearchProvider, buildRecentPeople, isValidEmail,
} from '../common/PeoplePicker';
import { toDateOnly, formatDateOnly } from '../../utils/dateUtils';
import { findDependencyCycle, workingDaysBetween } from '../../utils/scheduleUtils';
import * as strings from 'SmartGanttWebPartStrings';
import { formatString } from '../localeUtils';

interface ITaskPanelProps {
  isOpen: boolean;
  task: ITask | null;
  tasks: ITask[];
  project: IProject;
  knownPhases: string[];
  knownUsers: string[];
  onSave: (data: Partial<ITask>) => Promise<void>;
  onDismiss: () => void;
  /** Working calendar used for baseline variance (working days). Defaults to Mon–Fri. */
  calendar?: IWorkingCalendar;
  /** Directory search for the assignee picker. Without it the picker falls
   *  back to suggestions from previously used assignees only. */
  spService?: IPeopleSearchProvider;
  /** Every task in the project, unfiltered. Used for the predecessor list,
   *  cycle checks and recent-assignee suggestions; falls back to `tasks`. */
  allTasks?: ITask[];
  /** Initial values for a NEW task (ignored when editing) — e.g. a status
   *  preset from a Kanban column or dates from a drag on the Gantt. */
  defaults?: Partial<ITask>;
}

type TabId = 'basic' | 'details' | 'links';

const EMPTY: Partial<ITask> = {
  title: '',
  description: '',
  startDate: '',
  dueDate: '',
  status: 'Not Started',
  priority: 'Medium',
  assignedTo: '',
  assignedToEmail: '',
  percentComplete: 0,
  parentTaskId: null,
  dependencies: [],
  dependencyLinks: {},
  notes: '',
  color: '',
  isMilestone: false,
  phase: '',
};

const DEP_TYPES: DependencyType[] = ['FS', 'SS', 'FF', 'SF'];

const isDefaultLink = (l: IDependencyLink): boolean => l.type === 'FS' && l.lag === 0;

// Lag/lead in working days for one predecessor. Keeps its own text so a
// half-typed "-" doesn't get overwritten by the parsed number mid-edit.
const LagInput: React.FC<{ value: number; onChange: (n: number) => void; ariaLabel: string }> = ({ value, onChange, ariaLabel }) => {
  const [text, setText] = React.useState(String(value));
  React.useEffect(() => {
    if (parseInt(text, 10) !== value) setText(String(value));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value]);
  return (
    <TextField
      type="number"
      ariaLabel={ariaLabel}
      value={text}
      min={-999}
      max={999}
      styles={{ root: { width: 80 } }}
      onChange={(_, v) => {
        const t = v ?? '';
        setText(t);
        const n = parseInt(t, 10);
        if (!isNaN(n)) onChange(Math.max(-999, Math.min(999, n)));
      }}
      onBlur={() => { if (isNaN(parseInt(text, 10))) setText(String(value)); }}
    />
  );
};

export const TaskPanel: React.FC<ITaskPanelProps> = ({
  isOpen, task, tasks, project, knownPhases, knownUsers, onSave, onDismiss,
  calendar, spService, allTasks, defaults,
}) => {
  const isEdit = !!task;
  const pool = allTasks || tasks;
  const cal = calendar || DEFAULT_WORKING_CALENDAR;
  const [form, setForm] = React.useState<Partial<ITask>>(EMPTY);
  const [saving, setSaving] = React.useState(false);
  const [errors, setErrors] = React.useState<Record<string, string>>({});
  const [saveError, setSaveError] = React.useState('');
  const [dirty, setDirty] = React.useState(false);
  const [activeTab, setActiveTab] = React.useState<TabId>('basic');
  const tabRefs = React.useRef<Record<TabId, HTMLButtonElement | null>>({ basic: null, details: null, links: null });

  React.useEffect(() => {
    if (isOpen) {
      setForm(task
        ? { ...task, dependencyLinks: { ...(task.dependencyLinks || {}) } }
        : { ...EMPTY, dependencyLinks: {}, ...(defaults || {}) });
      setErrors({});
      setSaveError('');
      setSaving(false);
      setDirty(false);
      setActiveTab('basic');
    }
    // `defaults` is intentionally not a dependency: it only seeds the form on open.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isOpen, task]);

  const set = (field: keyof ITask, value: any): void => {
    setForm(prev => ({ ...prev, [field]: value }));
    setErrors(prev => ({ ...prev, [field]: '' }));
    setDirty(true);
  };

  // Assignee is stored as free text plus an optional email — the picker sets both.
  const setAssignee = (name: string, email: string): void => {
    setForm(prev => ({ ...prev, assignedTo: name, assignedToEmail: email }));
    setErrors(prev => ({ ...prev, assignee: '' }));
    setDirty(true);
  };

  const handleDismiss = (): void => {
    if (dirty && !saving && !window.confirm(strings.TaskPanel_DiscardUnsavedChangesConfirm)) return;
    onDismiss();
  };

  const validate = (): boolean => {
    const errs: Record<string, string> = {};
    if (!form.title?.trim()) errs.title = strings.TaskPanel_TaskNameRequired;
    // Normalize both sides to YYYY-MM-DD before comparing — an edited field
    // holds a date-only string while an untouched one may still be a full ISO
    // value, and comparing those directly gives false positives.
    const start = toDateOnly(form.startDate);
    const due = toDateOnly(form.dueDate);
    if (due && start && due < start) {
      errs.dueDate = strings.TaskPanel_DueDateError;
    }
    // The picker shows the message inline under its email field.
    if (!isValidEmail(form.assignedToEmail || '')) {
      errs.assignee = strings.Common_PeoplePicker_InvalidEmail;
    }
    setErrors(errs);
    return Object.keys(errs).length === 0;
  };

  const handleSave = async (): Promise<void> => {
    if (!validate()) {
      // The invalid email lives on the Basic tab.
      if (!form.title?.trim() || !isValidEmail(form.assignedToEmail || '')) setActiveTab('basic');
      return;
    }
    setSaving(true);
    setSaveError('');
    try {
      const deps = form.dependencies || [];
      // Only links for predecessors that are still present, and only
      // non-default ones (FS with no lag is what a missing entry means).
      const links: Record<number, IDependencyLink> = {};
      deps.forEach(id => {
        const l = form.dependencyLinks?.[id];
        if (l && !isDefaultLink(l)) links[id] = l;
      });
      await onSave({
        ...form,
        title: form.title!.trim(),
        assignedTo: (form.assignedTo || '').trim(),
        assignedToEmail: (form.assignedToEmail || '').trim(),
        startDate: toDateOnly(form.startDate),
        dueDate: toDateOnly(form.dueDate),
        dependencies: deps,
        dependencyLinks: links,
      });
    } catch (e) {
      setSaveError(e instanceof Error ? e.message : strings.TaskPanel_SaveFallbackError);
    } finally {
      setSaving(false);
    }
  };

  const statusLabels: Record<TaskStatus, string> = {
    'Not Started': strings.Status_NotStarted,
    'In Progress': strings.Status_InProgress,
    'Completed': strings.Status_Completed,
    'On Hold': strings.Status_OnHold,
    'Cancelled': strings.Status_Cancelled,
  };
  const priorityLabels: Record<TaskPriority, string> = {
    'Critical': strings.Priority_Critical,
    'High': strings.Priority_High,
    'Medium': strings.Priority_Medium,
    'Low': strings.Priority_Low,
  };
  const depTypeLabels: Record<DependencyType, string> = {
    FS: strings.Panel_Task_DepTypeFS,
    SS: strings.Panel_Task_DepTypeSS,
    FF: strings.Panel_Task_DepTypeFF,
    SF: strings.Panel_Task_DepTypeSF,
  };
  const statusOptions: IDropdownOption[] = TASK_STATUS_OPTIONS.map(s => ({ key: s, text: statusLabels[s] }));
  const priorityOptions: IDropdownOption[] = TASK_PRIORITY_OPTIONS.map(p => ({ key: p, text: priorityLabels[p] }));
  const depTypeOptions: IDropdownOption[] = DEP_TYPES.map(t => ({ key: t, text: depTypeLabels[t] }));

  // Directory matches come from spService; previously used assignees (with
  // the email last stored for them) are offered alongside.
  const recentPeople = React.useMemo(() => buildRecentPeople(knownUsers, pool), [knownUsers, pool]);

  // A task with sub-tasks of its own can't also become a sub-task itself —
  // the Gantt/List views only render one level of nesting, so a deeper chain
  // would make those children silently disappear.
  const hasChildren = !!task && pool.some(t => t.parentTaskId === task.id);

  // Parent task options — exclude self and already-children
  const parentOptions: IDropdownOption[] = [
    { key: '', text: strings.TaskPanel_ParentTaskNone },
    ...(hasChildren ? [] : pool
      .filter(t => t.id !== task?.id && !t.parentTaskId)
      .map(t => ({ key: t.id, text: t.title }))),
  ];

  // Tasks available to add as predecessors: not self, not already selected,
  // and not a task that (transitively) depends on this one — that link would
  // create a dependency cycle. A new task has no id yet, so nothing can cycle.
  const currentDeps = form.dependencies || [];
  const links = form.dependencyLinks || {};

  const addableDepOptions: IDropdownOption[] = [
    { key: '', text: strings.TaskPanel_SelectATaskOption },
    ...pool
      .filter(t => t.id !== task?.id
        && !currentDeps.includes(t.id)
        && !(task && findDependencyCycle(pool, t.id, task.id)))
      .map(t => ({ key: t.id, text: t.title })),
  ];

  const addDependency = (id: number): void => {
    if (!id || currentDeps.includes(id)) return;
    setForm(prev => ({ ...prev, dependencies: [...(prev.dependencies || []), id] }));
    setDirty(true);
  };

  const removeDependency = (id: number): void => {
    setForm(prev => {
      const nextLinks = { ...(prev.dependencyLinks || {}) };
      delete nextLinks[id];
      return { ...prev, dependencies: (prev.dependencies || []).filter(d => d !== id), dependencyLinks: nextLinks };
    });
    setDirty(true);
  };

  const updateLink = (id: number, patch: Partial<IDependencyLink>): void => {
    setForm(prev => {
      const current: IDependencyLink = prev.dependencyLinks?.[id] || { type: 'FS', lag: 0 };
      const next: IDependencyLink = { ...current, ...patch };
      const nextLinks = { ...(prev.dependencyLinks || {}) };
      if (isDefaultLink(next)) delete nextLinks[id]; else nextLinks[id] = next;
      return { ...prev, dependencyLinks: nextLinks };
    });
    setDirty(true);
  };

  // Baseline variance in working days (positive = later than planned).
  const varianceText = (baseline: string | undefined, current: string | undefined): string => {
    const b = toDateOnly(baseline);
    const c = toDateOnly(current);
    if (!b || !c) return '';
    const n = workingDaysBetween(b, c, cal);
    if (n === 0) return strings.Panel_Task_VarianceOnBaseline;
    return formatString(n > 0 ? strings.Panel_Task_VarianceLater : strings.Panel_Task_VarianceEarlier, { days: Math.abs(n) });
  };

  const statusColor = STATUS_COLORS[form.status as TaskStatus] || '#0078D4';

  const tabs: Array<{ id: TabId; label: string }> = [
    { id: 'basic', label: strings.TaskPanel_TabBasic },
    { id: 'details', label: strings.TaskPanel_TabDetails },
    { id: 'links', label: strings.TaskPanel_TabLinks },
  ];

  const handleTabKeyDown = (e: React.KeyboardEvent, idx: number): void => {
    let next = -1;
    if (e.key === 'ArrowRight') next = (idx + 1) % tabs.length;
    else if (e.key === 'ArrowLeft') next = (idx - 1 + tabs.length) % tabs.length;
    else if (e.key === 'Home') next = 0;
    else if (e.key === 'End') next = tabs.length - 1;
    if (next < 0) return;
    e.preventDefault();
    setActiveTab(tabs[next].id);
    tabRefs.current[tabs[next].id]?.focus();
  };

  const tabStyle = (tab: TabId): React.CSSProperties => ({
    padding: '7px 16px',
    fontSize: 13,
    fontWeight: activeTab === tab ? 600 : 400,
    color: activeTab === tab ? 'var(--themePrimary, #0078D4)' : 'var(--neutralSecondary, #605E5C)',
    borderBottom: activeTab === tab ? '2px solid var(--themePrimary, #0078D4)' : '2px solid transparent',
    background: 'transparent',
    border: 'none',
    cursor: 'pointer',
    transition: 'all 0.1s',
  });

  const hasBaseline = !!(task && (task.baselineStart || task.baselineDue));

  return (
    <Panel
      isOpen={isOpen}
      type={PanelType.medium}
      headerText={isEdit ? formatString(strings.TaskPanel_EditHeader, { taskTitle: task!.title }) : strings.TaskPanel_NewTaskHeader}
      onDismiss={handleDismiss}
      isFooterAtBottom
      onRenderFooterContent={() => (
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton
            disabled={saving}
            onClick={handleSave}
            style={{ minWidth: 120 }}
          >
            {saving && <Spinner size={SpinnerSize.small} style={{ marginRight: 6 }} />}
            {saving ? (isEdit ? strings.TaskPanel_SavingLabel : strings.TaskPanel_CreatingLabel) : (isEdit ? strings.TaskPanel_SaveChangesButton : strings.TaskPanel_CreateTaskButton)}
          </PrimaryButton>
          <DefaultButton text={strings.TaskPanel_CancelButton} onClick={handleDismiss} disabled={saving} />
        </Stack>
      )}
    >
      <div>
        {saveError && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={() => setSaveError('')}
            dismissButtonAriaLabel={strings.TaskPanel_DismissAriaLabel}
            styles={{ root: { marginTop: 8, marginBottom: 12 } }}
          >
            {saveError}
          </MessageBar>
        )}

        {/* Project context banner */}
        <div style={{
          display: 'flex', alignItems: 'center', gap: 8,
          background: `${project.color}12`,
          borderRadius: 4, padding: '8px 12px', marginBottom: 16,
          marginTop: 8,
        }}>
          <div style={{ width: 10, height: 10, borderRadius: '50%', background: project.color }} />
          <span style={{ fontSize: 12, color: 'var(--neutralSecondary, #605E5C)' }}>
            {strings.TaskPanel_ProjectLabelPrefix}<strong style={{ color: 'var(--neutralPrimary, #323130)' }}>{project.title}</strong>
          </span>
        </div>

        {/* Tabs */}
        <div role="tablist" style={{ display: 'flex', borderBottom: '1px solid var(--neutralLight, #EDEBE9)', marginBottom: 20 }}>
          {tabs.map((t, i) => (
            <button
              key={t.id}
              ref={el => { tabRefs.current[t.id] = el; }}
              id={`taskpanel-tab-${t.id}`}
              role="tab"
              type="button"
              aria-selected={activeTab === t.id}
              aria-controls={`taskpanel-tabpanel-${t.id}`}
              tabIndex={activeTab === t.id ? 0 : -1}
              style={tabStyle(t.id)}
              onClick={() => setActiveTab(t.id)}
              onKeyDown={e => handleTabKeyDown(e, i)}
            >
              {t.label}
            </button>
          ))}
        </div>

        {/* ── BASIC TAB ── */}
        {activeTab === 'basic' && (
          <div role="tabpanel" id="taskpanel-tabpanel-basic" aria-labelledby="taskpanel-tab-basic">
            <Stack tokens={{ childrenGap: 16 }}>
              <TextField
                label={strings.TaskPanel_TaskNameLabel}
                value={form.title || ''}
                onChange={(_, v) => set('title', v || '')}
                required
                errorMessage={errors.title}
                autoFocus
                placeholder={strings.TaskPanel_TaskNamePlaceholder}
              />

              <TextField
                label={strings.TaskPanel_DescriptionLabel}
                value={form.description || ''}
                onChange={(_, v) => set('description', v || '')}
                multiline
                rows={2}
                resizable={false}
                placeholder={strings.TaskPanel_DescriptionPlaceholder}
              />

              <Stack horizontal tokens={{ childrenGap: 12 }}>
                <Stack.Item grow>
                  <TextField
                    label={strings.TaskPanel_StartDateLabel}
                    type="date"
                    value={form.startDate ? form.startDate.split('T')[0] : ''}
                    onChange={(_, v) => set('startDate', v || '')}
                  />
                </Stack.Item>
                <Stack.Item grow>
                  <TextField
                    label={strings.TaskPanel_DueDateLabel}
                    type="date"
                    value={form.dueDate ? form.dueDate.split('T')[0] : ''}
                    onChange={(_, v) => { set('dueDate', v || ''); }}
                    errorMessage={errors.dueDate}
                  />
                </Stack.Item>
              </Stack>

              <Stack horizontal tokens={{ childrenGap: 12 }}>
                <Stack.Item grow>
                  <Dropdown
                    label={strings.TaskPanel_StatusLabel}
                    selectedKey={form.status}
                    options={statusOptions}
                    onChange={(_, opt) => {
                      if (!opt) return;
                      const s = opt.key as TaskStatus;
                      const updates: Partial<ITask> = { status: s };
                      if (s === 'Completed') updates.percentComplete = 100;
                      if (s === 'Not Started') updates.percentComplete = 0;
                      setForm(prev => ({ ...prev, ...updates }));
                      setDirty(true);
                    }}
                    onRenderOption={opt => opt ? (
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <div style={{
                          width: 8, height: 8, borderRadius: '50%',
                          background: STATUS_COLORS[opt.key as TaskStatus] || '#8B929A',
                          flexShrink: 0,
                        }} />
                        {opt.text}
                      </div>
                    ) : null}
                  />
                </Stack.Item>
                <Stack.Item grow>
                  <Dropdown
                    label={strings.TaskPanel_PriorityLabel}
                    selectedKey={form.priority}
                    options={priorityOptions}
                    onChange={(_, opt) => opt && set('priority', opt.key)}
                    onRenderOption={opt => opt ? (
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <div style={{
                          width: 8, height: 8, borderRadius: '50%',
                          background: PRIORITY_COLORS[opt.key as TaskPriority] || '#0078D4',
                          flexShrink: 0,
                        }} />
                        {opt.text}
                      </div>
                    ) : null}
                  />
                </Stack.Item>
              </Stack>

              {/* Progress */}
              <div>
                <Label>{formatString(strings.TaskPanel_PercentCompleteLabel, { percent: `‹PCT›` }).split('‹PCT›').map((part, i, arr) => (
                  <React.Fragment key={i}>
                    {part}
                    {i < arr.length - 1 && <strong style={{ color: statusColor }}>{form.percentComplete}</strong>}
                  </React.Fragment>
                ))}</Label>
                <Slider
                  min={0}
                  max={100}
                  step={5}
                  value={form.percentComplete || 0}
                  onChange={v => {
                    // Keep status roughly in sync with progress — otherwise a
                    // task can end up "Completed" at 50% or "Not Started" at 100%
                    // with nothing flagging the mismatch.
                    setForm(prev => {
                      const updates: Partial<ITask> = { percentComplete: v };
                      if (v === 100 && prev.status !== 'Cancelled') updates.status = 'Completed';
                      else if (v > 0 && v < 100 && (prev.status === 'Not Started' || prev.status === 'Completed')) updates.status = 'In Progress';
                      else if (v === 0 && prev.status === 'Completed') updates.status = 'Not Started';
                      return { ...prev, ...updates };
                    });
                    setDirty(true);
                  }}
                  showValue={false}
                  styles={{
                    activeSection: { background: statusColor },
                    thumb: { borderColor: statusColor },
                  }}
                />
              </div>

              <PeoplePicker
                key={task ? `t${task.id}` : 'new'}
                label={strings.TaskPanel_AssignedToLabel}
                name={form.assignedTo || ''}
                email={form.assignedToEmail || ''}
                onChange={setAssignee}
                recent={recentPeople}
                spService={spService}
                placeholder={strings.TaskPanel_AssignedToPlaceholder}
              />
            </Stack>
          </div>
        )}

        {/* ── DETAILS TAB ── */}
        {activeTab === 'details' && (
          <div role="tabpanel" id="taskpanel-tabpanel-details" aria-labelledby="taskpanel-tab-details">
            <Stack tokens={{ childrenGap: 16 }}>
              <AutocompleteField
                label={strings.TaskPanel_PhaseLabel}
                value={form.phase || ''}
                suggestions={knownPhases}
                onChange={v => set('phase', v)}
                placeholder={strings.TaskPanel_PhasePlaceholder}
              />
              <div style={{ fontSize: 11, color: 'var(--neutralSecondary, #605E5C)', marginTop: 4 }}>
                {strings.TaskPanel_PhaseHint}
              </div>

              <Toggle
                label={strings.TaskPanel_MilestoneLabel}
                checked={form.isMilestone || false}
                onChange={(_, v) => set('isMilestone', v)}
                onText={strings.TaskPanel_MilestoneOnText}
                offText={strings.TaskPanel_MilestoneOffText}
              />

              {/* Bar color override */}
              <div>
                <Label>{strings.TaskPanel_CustomBarColorLabel} <span style={{ color: 'var(--neutralSecondary, #605E5C)', fontWeight: 400 }}>{strings.TaskPanel_CustomBarColorOptionalSuffix}</span></Label>
                <ColorSwatchPicker
                  colors={PROJECT_COLORS}
                  value={form.color || ''}
                  onChange={c => set('color', c)}
                  allowAuto
                  ariaLabel={strings.TaskPanel_CustomBarColorLabel}
                />
                {!form.color && (
                  <div style={{ fontSize: 11, color: 'var(--neutralSecondary, #605E5C)', marginTop: 4 }}>
                    {strings.TaskPanel_AutoColorHint}
                  </div>
                )}
              </div>

              <TextField
                label={strings.TaskPanel_NotesLabel}
                value={form.notes || ''}
                onChange={(_, v) => set('notes', v || '')}
                multiline
                rows={5}
                resizable={false}
                placeholder={strings.TaskPanel_NotesPlaceholder}
              />
            </Stack>
          </div>
        )}

        {/* ── LINKS TAB ── */}
        {activeTab === 'links' && (
          <div role="tabpanel" id="taskpanel-tabpanel-links" aria-labelledby="taskpanel-tab-links">
            <Stack tokens={{ childrenGap: 16 }}>
              <div>
                <Dropdown
                  label={strings.TaskPanel_ParentTaskLabel}
                  selectedKey={form.parentTaskId ?? ''}
                  options={parentOptions}
                  onChange={(_, opt) => set('parentTaskId', opt?.key || null)}
                />
                <div style={{ fontSize: 11, color: 'var(--neutralSecondary, #605E5C)', marginTop: 4 }}>
                  {hasChildren
                    ? strings.TaskPanel_ParentTaskHasChildrenHint
                    : strings.TaskPanel_ParentTaskHint}
                </div>
              </div>

              {pool.filter(t => t.id !== task?.id).length > 0 && (
                <div>
                  <Label>{strings.TaskPanel_DependsOnLabel}</Label>

                  {/* Current predecessors: link type + lag per row */}
                  {currentDeps.length > 0 && (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 10 }}>
                      {currentDeps.map(depId => {
                        const dep = pool.find(t => t.id === depId);
                        if (!dep) return null;
                        const link: IDependencyLink = links[depId] || { type: 'FS', lag: 0 };
                        return (
                          <div
                            key={depId}
                            role="group"
                            aria-label={dep.title}
                            style={{
                              display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap',
                              padding: '6px 8px',
                              background: 'var(--themeLighterAlt, #EFF6FC)',
                              border: '1px solid var(--themeLight, #90C8F6)',
                              borderRadius: 6,
                            }}
                          >
                            <span
                              style={{
                                display: 'inline-flex', alignItems: 'center', gap: 6, flex: '1 1 140px', minWidth: 0,
                                fontSize: 12, color: 'var(--neutralPrimary, #323130)',
                              }}
                            >
                              <span
                                aria-hidden="true"
                                style={{
                                  width: 6, height: 6, borderRadius: '50%', flexShrink: 0,
                                  background: STATUS_COLORS[dep.status],
                                }}
                              />
                              <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={dep.title}>{dep.title}</span>
                            </span>
                            <Dropdown
                              ariaLabel={formatString(strings.Panel_Task_LinkTypeAriaLabel, { taskTitle: dep.title })}
                              selectedKey={link.type}
                              options={depTypeOptions}
                              onChange={(_, opt) => { if (opt) updateLink(depId, { type: opt.key as DependencyType }); }}
                              styles={{ root: { width: 190 } }}
                            />
                            <LagInput
                              value={link.lag}
                              ariaLabel={formatString(strings.Panel_Task_LagAriaLabel, { taskTitle: dep.title })}
                              onChange={n => updateLink(depId, { lag: n })}
                            />
                            <button
                              type="button"
                              onClick={() => removeDependency(depId)}
                              style={{
                                background: 'none', border: 'none', cursor: 'pointer',
                                padding: '0 4px', color: 'var(--themePrimary, #0078D4)', fontSize: 18,
                                lineHeight: 1, display: 'flex', alignItems: 'center',
                              }}
                              title={formatString(strings.TaskPanel_RemoveDependencyTitle, { taskTitle: dep.title })}
                              aria-label={formatString(strings.TaskPanel_RemoveDependencyTitle, { taskTitle: dep.title })}
                            >
                              ×
                            </button>
                          </div>
                        );
                      })}
                      <div style={{ fontSize: 11, color: 'var(--neutralSecondary, #605E5C)' }}>
                        {strings.Panel_Task_LagHint}
                      </div>
                    </div>
                  )}

                  {/* Add a new dependency */}
                  {addableDepOptions.length > 1 && (
                    <Dropdown
                      placeholder={strings.TaskPanel_AddDependencyPlaceholder}
                      ariaLabel={strings.TaskPanel_AddDependencyPlaceholder}
                      selectedKey={''}
                      options={addableDepOptions}
                      onChange={(_, opt) => {
                        if (opt?.key) addDependency(opt.key as number);
                      }}
                    />
                  )}

                  <div style={{ fontSize: 11, color: 'var(--neutralSecondary, #605E5C)', marginTop: 6 }}>
                    {strings.TaskPanel_DependsOnHint}
                  </div>
                </div>
              )}

              {/* Baseline (read-only) */}
              {hasBaseline && task && (
                <div>
                  <Label>{strings.Panel_Task_BaselineHeader}</Label>
                  <div style={{ fontSize: 12, color: 'var(--neutralPrimary, #323130)', display: 'flex', flexDirection: 'column', gap: 4 }}>
                    {task.baselineStart && (
                      <div>
                        <strong>{strings.Panel_Task_BaselineStartLabel}</strong>{' '}
                        {formatDateOnly(task.baselineStart, 'MMM d, yyyy')}
                        {varianceText(task.baselineStart, form.startDate) && (
                          <span style={{ color: 'var(--neutralSecondary, #605E5C)' }}> ({varianceText(task.baselineStart, form.startDate)})</span>
                        )}
                      </div>
                    )}
                    {task.baselineDue && (
                      <div>
                        <strong>{strings.Panel_Task_BaselineDueLabel}</strong>{' '}
                        {formatDateOnly(task.baselineDue, 'MMM d, yyyy')}
                        {varianceText(task.baselineDue, form.dueDate) && (
                          <span style={{ color: 'var(--neutralSecondary, #605E5C)' }}> ({varianceText(task.baselineDue, form.dueDate)})</span>
                        )}
                      </div>
                    )}
                  </div>
                </div>
              )}
            </Stack>
          </div>
        )}
      </div>
    </Panel>
  );
};

export default TaskPanel;
