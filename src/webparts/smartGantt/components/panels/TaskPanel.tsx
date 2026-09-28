import * as React from 'react';
import {
  Panel, PanelType, TextField, Dropdown, IDropdownOption,
  PrimaryButton, DefaultButton, Stack, Label, Toggle, Spinner, SpinnerSize,
  Slider, MessageBar, MessageBarType,
} from '@fluentui/react';
import {
  ITask, IProject, TaskStatus, TaskPriority,
  TASK_STATUS_OPTIONS, TASK_PRIORITY_OPTIONS,
  STATUS_COLORS, PRIORITY_COLORS, PROJECT_COLORS,
} from '../../models';
import { AutocompleteField } from '../common/AutocompleteField';
import { ColorSwatchPicker } from '../common/ColorSwatchPicker';
import { toDateOnly } from '../../utils/dateUtils';
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
}

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
  notes: '',
  color: '',
  isMilestone: false,
  phase: '',
};

export const TaskPanel: React.FC<ITaskPanelProps> = ({
  isOpen, task, tasks, project, knownPhases, knownUsers, onSave, onDismiss,
}) => {
  const isEdit = !!task;
  const [form, setForm] = React.useState<Partial<ITask>>(EMPTY);
  const [saving, setSaving] = React.useState(false);
  const [errors, setErrors] = React.useState<Record<string, string>>({});
  const [saveError, setSaveError] = React.useState('');
  const [dirty, setDirty] = React.useState(false);
  const [activeTab, setActiveTab] = React.useState<'basic' | 'details' | 'links'>('basic');

  React.useEffect(() => {
    if (isOpen) {
      setForm(task ? { ...task } : { ...EMPTY });
      setErrors({});
      setSaveError('');
      setSaving(false);
      setDirty(false);
      setActiveTab('basic');
    }
  }, [isOpen, task]);

  const set = (field: keyof ITask, value: any): void => {
    setForm(prev => ({ ...prev, [field]: value }));
    setErrors(prev => ({ ...prev, [field]: '' }));
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
    setErrors(errs);
    return Object.keys(errs).length === 0;
  };

  const handleSave = async (): Promise<void> => {
    if (!validate()) return;
    setSaving(true);
    setSaveError('');
    try {
      await onSave({
        ...form,
        title: form.title!.trim(),
        startDate: toDateOnly(form.startDate),
        dueDate: toDateOnly(form.dueDate),
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
  const statusOptions: IDropdownOption[] = TASK_STATUS_OPTIONS.map(s => ({ key: s, text: statusLabels[s] }));
  const priorityOptions: IDropdownOption[] = TASK_PRIORITY_OPTIONS.map(p => ({ key: p, text: priorityLabels[p] }));

  // A task with sub-tasks of its own can't also become a sub-task itself —
  // the Gantt/List views only render one level of nesting, so a deeper chain
  // would make those children silently disappear.
  const hasChildren = !!task && tasks.some(t => t.parentTaskId === task.id);

  // Parent task options — exclude self and already-children
  const parentOptions: IDropdownOption[] = [
    { key: '', text: strings.TaskPanel_ParentTaskNone },
    ...(hasChildren ? [] : tasks
      .filter(t => t.id !== task?.id && !t.parentTaskId)
      .map(t => ({ key: t.id, text: t.title }))),
  ];

  // Tasks available to add as dependencies: not self, not already selected,
  // and not a task that (transitively) depends on this one — that would
  // create a dependency cycle.
  const currentDeps = form.dependencies || [];
  const dependentIds = React.useMemo(() => {
    const set = new Set<number>();
    if (!task) return set;
    const dependentsOf = new Map<number, number[]>();
    tasks.forEach(t => {
      t.dependencies.forEach(depId => {
        if (!dependentsOf.has(depId)) dependentsOf.set(depId, []);
        dependentsOf.get(depId)!.push(t.id);
      });
    });
    const queue = [task.id];
    while (queue.length) {
      const cur = queue.pop()!;
      for (const next of dependentsOf.get(cur) || []) {
        if (!set.has(next)) { set.add(next); queue.push(next); }
      }
    }
    return set;
  }, [tasks, task]);

  const addableDepOptions: IDropdownOption[] = [
    { key: '', text: strings.TaskPanel_SelectATaskOption },
    ...tasks
      .filter(t => t.id !== task?.id && !currentDeps.includes(t.id) && !dependentIds.has(t.id))
      .map(t => ({ key: t.id, text: t.title })),
  ];

  const addDependency = (id: number): void => {
    if (!id || currentDeps.includes(id)) return;
    setForm(prev => ({ ...prev, dependencies: [...(prev.dependencies || []), id] }));
  };

  const removeDependency = (id: number): void => {
    setForm(prev => ({ ...prev, dependencies: (prev.dependencies || []).filter(d => d !== id) }));
  };

  const statusColor = STATUS_COLORS[form.status as TaskStatus] || '#0078D4';

  const tabStyle = (tab: 'basic' | 'details' | 'links'): React.CSSProperties => ({
    padding: '7px 16px',
    fontSize: 13,
    fontWeight: activeTab === tab ? 600 : 400,
    color: activeTab === tab ? '#0078D4' : '#605E5C',
    borderBottom: activeTab === tab ? '2px solid #0078D4' : '2px solid transparent',
    background: 'transparent',
    border: 'none',
    cursor: 'pointer',
    transition: 'all 0.1s',
  });

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
          <span style={{ fontSize: 12, color: '#605E5C' }}>
            {strings.TaskPanel_ProjectLabelPrefix}<strong style={{ color: '#323130' }}>{project.title}</strong>
          </span>
        </div>

        {/* Tabs */}
        <div style={{ display: 'flex', borderBottom: '1px solid #EDEBE9', marginBottom: 20 }}>
          <button style={tabStyle('basic')} onClick={() => setActiveTab('basic')}>{strings.TaskPanel_TabBasic}</button>
          <button style={tabStyle('details')} onClick={() => setActiveTab('details')}>{strings.TaskPanel_TabDetails}</button>
          <button style={tabStyle('links')} onClick={() => setActiveTab('links')}>{strings.TaskPanel_TabLinks}</button>
        </div>

        {/* ── BASIC TAB ── */}
        {activeTab === 'basic' && (
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

            <AutocompleteField
              label={strings.TaskPanel_AssignedToLabel}
              value={form.assignedTo || ''}
              suggestions={knownUsers}
              onChange={v => set('assignedTo', v)}
              placeholder={strings.TaskPanel_AssignedToPlaceholder}
            />
          </Stack>
        )}

        {/* ── DETAILS TAB ── */}
        {activeTab === 'details' && (
          <Stack tokens={{ childrenGap: 16 }}>
            <AutocompleteField
              label={strings.TaskPanel_PhaseLabel}
              value={form.phase || ''}
              suggestions={knownPhases}
              onChange={v => set('phase', v)}
              placeholder={strings.TaskPanel_PhasePlaceholder}
            />
            <div style={{ fontSize: 11, color: '#605E5C', marginTop: 4 }}>
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
              <Label>{strings.TaskPanel_CustomBarColorLabel} <span style={{ color: '#605E5C', fontWeight: 400 }}>{strings.TaskPanel_CustomBarColorOptionalSuffix}</span></Label>
              <ColorSwatchPicker
                colors={PROJECT_COLORS}
                value={form.color || ''}
                onChange={c => set('color', c)}
                allowAuto
              />
              {!form.color && (
                <div style={{ fontSize: 11, color: '#605E5C', marginTop: 4 }}>
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
        )}

        {/* ── LINKS TAB ── */}
        {activeTab === 'links' && (
          <Stack tokens={{ childrenGap: 16 }}>
            <div>
              <Dropdown
                label={strings.TaskPanel_ParentTaskLabel}
                selectedKey={form.parentTaskId ?? ''}
                options={parentOptions}
                onChange={(_, opt) => set('parentTaskId', opt?.key || null)}
              />
              <div style={{ fontSize: 11, color: '#605E5C', marginTop: 4 }}>
                {hasChildren
                  ? strings.TaskPanel_ParentTaskHasChildrenHint
                  : strings.TaskPanel_ParentTaskHint}
              </div>
            </div>

            {tasks.filter(t => t.id !== task?.id).length > 0 && (
              <div>
                <Label>{strings.TaskPanel_DependsOnLabel}</Label>

                {/* Current dependencies as removable chips */}
                {currentDeps.length > 0 && (
                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: 8 }}>
                    {currentDeps.map(depId => {
                      const dep = tasks.find(t => t.id === depId);
                      if (!dep) return null;
                      return (
                        <span
                          key={depId}
                          style={{
                            display: 'inline-flex', alignItems: 'center', gap: 6,
                            padding: '3px 8px 3px 10px',
                            background: '#EFF6FC', border: '1px solid #90C8F6',
                            borderRadius: 12, fontSize: 12, color: '#0078D4',
                          }}
                        >
                          <span
                            style={{
                              width: 6, height: 6, borderRadius: '50%', flexShrink: 0,
                              background: STATUS_COLORS[dep.status],
                            }}
                          />
                          {dep.title}
                          <button
                            onClick={() => removeDependency(depId)}
                            style={{
                              background: 'none', border: 'none', cursor: 'pointer',
                              padding: '0 2px', color: '#0078D4', fontSize: 14,
                              lineHeight: 1, display: 'flex', alignItems: 'center',
                            }}
                            title={formatString(strings.TaskPanel_RemoveDependencyTitle, { taskTitle: dep.title })}
                          >
                            ×
                          </button>
                        </span>
                      );
                    })}
                  </div>
                )}

                {/* Add a new dependency */}
                {addableDepOptions.length > 1 && (
                  <Dropdown
                    placeholder={strings.TaskPanel_AddDependencyPlaceholder}
                    selectedKey={''}
                    options={addableDepOptions}
                    onChange={(_, opt) => {
                      if (opt?.key) addDependency(opt.key as number);
                    }}
                  />
                )}

                <div style={{ fontSize: 11, color: '#605E5C', marginTop: 6 }}>
                  {strings.TaskPanel_DependsOnHint}
                </div>
              </div>
            )}
          </Stack>
        )}
      </div>
    </Panel>
  );
};

export default TaskPanel;
