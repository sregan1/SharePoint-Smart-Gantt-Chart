import * as React from 'react';
import {
  Panel, PanelType, TextField, Dropdown, IDropdownOption,
  PrimaryButton, DefaultButton, Stack, Label, Toggle, Spinner, SpinnerSize,
  Slider,
} from '@fluentui/react';
import {
  ITask, IProject, TaskStatus, TaskPriority,
  TASK_STATUS_OPTIONS, TASK_PRIORITY_OPTIONS,
  STATUS_COLORS, PRIORITY_COLORS, PROJECT_COLORS,
} from '../../models';
import { AutocompleteField } from '../common/AutocompleteField';

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
  const [activeTab, setActiveTab] = React.useState<'basic' | 'details' | 'links'>('basic');

  React.useEffect(() => {
    if (isOpen) {
      setForm(task ? { ...task } : { ...EMPTY });
      setErrors({});
      setSaving(false);
      setActiveTab('basic');
    }
  }, [isOpen, task]);

  const set = (field: keyof ITask, value: any): void => {
    setForm(prev => ({ ...prev, [field]: value }));
    setErrors(prev => ({ ...prev, [field]: '' }));
  };

  const validate = (): boolean => {
    const errs: Record<string, string> = {};
    if (!form.title?.trim()) errs.title = 'Task name is required.';
    if (form.dueDate && form.startDate && form.dueDate < form.startDate) {
      errs.dueDate = 'Due date must be on or after start date.';
    }
    setErrors(errs);
    return Object.keys(errs).length === 0;
  };

  const handleSave = async (): Promise<void> => {
    if (!validate()) return;
    setSaving(true);
    try {
      await onSave({
        ...form,
        title: form.title!.trim(),
        startDate: form.startDate ? new Date(form.startDate).toISOString() : '',
        dueDate: form.dueDate ? new Date(form.dueDate).toISOString() : '',
      });
    } finally {
      setSaving(false);
    }
  };

  const statusOptions: IDropdownOption[] = TASK_STATUS_OPTIONS.map(s => ({ key: s, text: s }));
  const priorityOptions: IDropdownOption[] = TASK_PRIORITY_OPTIONS.map(p => ({ key: p, text: p }));

  // Parent task options — exclude self and already-children
  const parentOptions: IDropdownOption[] = [
    { key: '', text: 'None (top-level task)' },
    ...tasks
      .filter(t => t.id !== task?.id && !t.parentTaskId)
      .map(t => ({ key: t.id, text: t.title })),
  ];

  // Tasks available to add as dependencies (not self, not already selected)
  const currentDeps = form.dependencies || [];
  const addableDepOptions: IDropdownOption[] = [
    { key: '', text: 'Select a task…' },
    ...tasks
      .filter(t => t.id !== task?.id && !currentDeps.includes(t.id))
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
      headerText={isEdit ? `Edit: ${task!.title}` : 'New Task'}
      onDismiss={onDismiss}
      isFooterAtBottom
      onRenderFooterContent={() => (
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton
            disabled={saving}
            onClick={handleSave}
            style={{ minWidth: 120 }}
          >
            {saving && <Spinner size={SpinnerSize.small} style={{ marginRight: 6 }} />}
            {saving ? (isEdit ? 'Saving…' : 'Creating…') : (isEdit ? 'Save Changes' : 'Create Task')}
          </PrimaryButton>
          <DefaultButton text="Cancel" onClick={onDismiss} disabled={saving} />
        </Stack>
      )}
    >
      <div>
        {/* Project context banner */}
        <div style={{
          display: 'flex', alignItems: 'center', gap: 8,
          background: `${project.color}12`,
          borderRadius: 4, padding: '8px 12px', marginBottom: 16,
          marginTop: 8,
        }}>
          <div style={{ width: 10, height: 10, borderRadius: '50%', background: project.color }} />
          <span style={{ fontSize: 12, color: '#605E5C' }}>
            Project: <strong style={{ color: '#323130' }}>{project.title}</strong>
          </span>
        </div>

        {/* Tabs */}
        <div style={{ display: 'flex', borderBottom: '1px solid #EDEBE9', marginBottom: 20 }}>
          <button style={tabStyle('basic')} onClick={() => setActiveTab('basic')}>Basic</button>
          <button style={tabStyle('details')} onClick={() => setActiveTab('details')}>Details</button>
          <button style={tabStyle('links')} onClick={() => setActiveTab('links')}>Links</button>
        </div>

        {/* ── BASIC TAB ── */}
        {activeTab === 'basic' && (
          <Stack tokens={{ childrenGap: 16 }}>
            <TextField
              label="Task Name"
              value={form.title || ''}
              onChange={(_, v) => set('title', v || '')}
              required
              errorMessage={errors.title}
              autoFocus
              placeholder="What needs to be done?"
            />

            <TextField
              label="Description"
              value={form.description || ''}
              onChange={(_, v) => set('description', v || '')}
              multiline
              rows={2}
              resizable={false}
              placeholder="Optional task description…"
            />

            <Stack horizontal tokens={{ childrenGap: 12 }}>
              <Stack.Item grow>
                <TextField
                  label="Start Date"
                  type="date"
                  value={form.startDate ? form.startDate.split('T')[0] : ''}
                  onChange={(_, v) => set('startDate', v || '')}
                />
              </Stack.Item>
              <Stack.Item grow>
                <TextField
                  label="Due Date"
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
                  label="Status"
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
                  label="Priority"
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
              <Label>% Complete: <strong style={{ color: statusColor }}>{form.percentComplete}%</strong></Label>
              <Slider
                min={0}
                max={100}
                step={5}
                value={form.percentComplete || 0}
                onChange={v => set('percentComplete', v)}
                showValue={false}
                styles={{
                  activeSection: { background: statusColor },
                  thumb: { borderColor: statusColor },
                }}
              />
            </div>

            <AutocompleteField
              label="Assigned To"
              value={form.assignedTo || ''}
              suggestions={knownUsers}
              onChange={v => set('assignedTo', v)}
              placeholder="Start typing a name…"
            />
          </Stack>
        )}

        {/* ── DETAILS TAB ── */}
        {activeTab === 'details' && (
          <Stack tokens={{ childrenGap: 16 }}>
            <AutocompleteField
              label="Phase"
              value={form.phase || ''}
              suggestions={knownPhases}
              onChange={v => set('phase', v)}
              placeholder="e.g. Discovery, Design, Development"
            />
            <div style={{ fontSize: 11, color: '#605E5C', marginTop: 4 }}>
              Groups tasks visually on the Gantt. Start typing to see existing phases.
            </div>

            <Toggle
              label="Milestone"
              checked={form.isMilestone || false}
              onChange={(_, v) => set('isMilestone', v)}
              onText="Yes — shown as ◆ on Gantt"
              offText="No"
            />

            {/* Bar color override */}
            <div>
              <Label>Custom Bar Color <span style={{ color: '#605E5C', fontWeight: 400 }}>(optional)</span></Label>
              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 6, alignItems: 'center' }}>
                <div
                  onClick={() => set('color', '')}
                  style={{
                    width: 28, height: 28, borderRadius: '50%',
                    background: 'linear-gradient(135deg, #ccc 50%, #fff 50%)',
                    cursor: 'pointer',
                    border: !form.color ? '3px solid #323130' : '2px solid #EDEBE9',
                    boxSizing: 'border-box',
                  }}
                  title="Auto color"
                />
                {PROJECT_COLORS.map(c => (
                  <div
                    key={c}
                    onClick={() => set('color', c)}
                    style={{
                      width: 28, height: 28, borderRadius: '50%',
                      background: c, cursor: 'pointer',
                      border: form.color === c ? '3px solid #323130' : '2px solid transparent',
                      outline: form.color === c ? `2px solid ${c}` : 'none',
                      outlineOffset: 2,
                      boxSizing: 'border-box',
                    }}
                  />
                ))}
                {/* Custom color swatch */}
                <label
                  title="Pick a custom color"
                  style={{ position: 'relative', width: 28, height: 28, cursor: 'pointer', flexShrink: 0 }}
                >
                  <div style={{
                    width: 28, height: 28, borderRadius: '50%',
                    background: form.color && !PROJECT_COLORS.includes(form.color)
                      ? form.color
                      : 'conic-gradient(red,yellow,lime,cyan,blue,magenta,red)',
                    border: form.color && !PROJECT_COLORS.includes(form.color)
                      ? '3px solid #323130' : '2px solid #EDEBE9',
                    outline: form.color && !PROJECT_COLORS.includes(form.color)
                      ? `2px solid ${form.color}` : 'none',
                    outlineOffset: 2, boxSizing: 'border-box',
                  }} />
                  <input
                    type="color"
                    value={form.color || '#0078D4'}
                    onChange={e => set('color', e.target.value)}
                    style={{ position: 'absolute', opacity: 0, width: 0, height: 0, pointerEvents: 'none' }}
                  />
                </label>
              </div>
              {!form.color && (
                <div style={{ fontSize: 11, color: '#605E5C', marginTop: 4 }}>
                  Auto: color follows Display Settings
                </div>
              )}
            </div>

            <TextField
              label="Notes"
              value={form.notes || ''}
              onChange={(_, v) => set('notes', v || '')}
              multiline
              rows={5}
              resizable={false}
              placeholder="Additional notes, links, context…"
            />
          </Stack>
        )}

        {/* ── LINKS TAB ── */}
        {activeTab === 'links' && (
          <Stack tokens={{ childrenGap: 16 }}>
            <div>
              <Dropdown
                label="Parent Task"
                selectedKey={form.parentTaskId ?? ''}
                options={parentOptions}
                onChange={(_, opt) => set('parentTaskId', opt?.key || null)}
              />
              <div style={{ fontSize: 11, color: '#605E5C', marginTop: 4 }}>
                Makes this a sub-task, shown indented below the parent.
              </div>
            </div>

            {tasks.filter(t => t.id !== task?.id).length > 0 && (
              <div>
                <Label>Depends On</Label>

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
                            title={`Remove dependency on "${dep.title}"`}
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
                    placeholder="Add a dependency…"
                    selectedKey={''}
                    options={addableDepOptions}
                    onChange={(_, opt) => {
                      if (opt?.key) addDependency(opt.key as number);
                    }}
                  />
                )}

                <div style={{ fontSize: 11, color: '#605E5C', marginTop: 6 }}>
                  This task cannot start until all dependencies are complete.
                  Arrows are drawn on the Gantt chart.
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
