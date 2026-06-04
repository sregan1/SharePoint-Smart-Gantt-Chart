import * as React from 'react';
import {
  Panel, PanelType, TextField, Dropdown, IDropdownOption,
  PrimaryButton, DefaultButton, Stack, Label, Spinner, SpinnerSize,
} from '@fluentui/react';
import { IProject, PROJECT_COLORS, PROJECT_STATUS_OPTIONS, ProjectStatus } from '../../models';

interface IProjectPanelProps {
  isOpen: boolean;
  project: IProject | null;
  onSave: (data: Partial<IProject>) => Promise<void>;
  onDismiss: () => void;
}

export const ProjectPanel: React.FC<IProjectPanelProps> = ({ isOpen, project, onSave, onDismiss }) => {
  const isEdit = !!project;

  const [title, setTitle] = React.useState('');
  const [description, setDescription] = React.useState('');
  const [color, setColor] = React.useState(PROJECT_COLORS[0]);
  const [startDate, setStartDate] = React.useState('');
  const [dueDate, setDueDate] = React.useState('');
  const [status, setStatus] = React.useState<ProjectStatus>('Active');
  const [saving, setSaving] = React.useState(false);
  const [errors, setErrors] = React.useState<Record<string, string>>({});

  // Populate form when editing
  React.useEffect(() => {
    if (isOpen) {
      if (project) {
        setTitle(project.title);
        setDescription(project.description);
        setColor(project.color || PROJECT_COLORS[0]);
        setStartDate(project.startDate ? project.startDate.split('T')[0] : '');
        setDueDate(project.dueDate ? project.dueDate.split('T')[0] : '');
        setStatus(project.status);
      } else {
        setTitle('');
        setDescription('');
        setColor(PROJECT_COLORS[0]);
        setStartDate('');
        setDueDate('');
        setStatus('Active');
      }
      setErrors({});
      setSaving(false);
    }
  }, [isOpen, project]);

  const validate = (): boolean => {
    const errs: Record<string, string> = {};
    if (!title.trim()) errs.title = 'Project name is required.';
    if (dueDate && startDate && dueDate < startDate) errs.dueDate = 'Due date must be after start date.';
    setErrors(errs);
    return Object.keys(errs).length === 0;
  };

  const handleSave = async (): Promise<void> => {
    if (!validate()) return;
    setSaving(true);
    try {
      await onSave({
        title: title.trim(),
        description,
        color,
        startDate: startDate ? new Date(startDate).toISOString() : '',
        dueDate: dueDate ? new Date(dueDate).toISOString() : '',
        status,
      });
    } finally {
      setSaving(false);
    }
  };

  const statusOptions: IDropdownOption[] = PROJECT_STATUS_OPTIONS.map(s => ({ key: s, text: s }));

  return (
    <Panel
      isOpen={isOpen}
      type={PanelType.smallFixedFar}
      headerText={isEdit ? 'Edit Project' : 'New Project'}
      onDismiss={onDismiss}
      isFooterAtBottom
      onRenderFooterContent={() => (
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton
            text={saving ? '' : isEdit ? 'Save Changes' : 'Create Project'}
            onClick={handleSave}
            disabled={saving}
          >
            {saving && <Spinner size={SpinnerSize.small} style={{ marginRight: 6 }} />}
            {saving ? (isEdit ? 'Saving…' : 'Creating…') : undefined}
          </PrimaryButton>
          <DefaultButton text="Cancel" onClick={onDismiss} disabled={saving} />
        </Stack>
      )}
    >
      <Stack tokens={{ childrenGap: 16 }} style={{ marginTop: 20 }}>
        {/* Project name */}
        <TextField
          label="Project Name"
          value={title}
          onChange={(_, v) => { setTitle(v || ''); setErrors(p => ({ ...p, title: '' })); }}
          required
          errorMessage={errors.title}
          autoFocus
          placeholder="e.g. Website Redesign"
        />

        {/* Description */}
        <TextField
          label="Description"
          value={description}
          onChange={(_, v) => setDescription(v || '')}
          multiline
          rows={3}
          placeholder="What is this project about?"
          resizable={false}
        />

        {/* Color picker */}
        <div>
          <Label>Project Color</Label>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 6, alignItems: 'center' }}>
            {PROJECT_COLORS.map(c => (
              <div
                key={c}
                onClick={() => setColor(c)}
                style={{
                  width: 30, height: 30, borderRadius: '50%', background: c,
                  cursor: 'pointer',
                  border: color === c ? '3px solid #323130' : '3px solid transparent',
                  outline: color === c ? `2px solid ${c}` : 'none',
                  outlineOffset: 2, transition: 'all 0.15s', boxSizing: 'border-box',
                }}
              />
            ))}
            {/* Custom color swatch */}
            <label
              title="Pick a custom color"
              style={{ position: 'relative', width: 30, height: 30, cursor: 'pointer', flexShrink: 0 }}
            >
              <div style={{
                width: 30, height: 30, borderRadius: '50%',
                background: !PROJECT_COLORS.includes(color) ? color : 'conic-gradient(red,yellow,lime,cyan,blue,magenta,red)',
                border: !PROJECT_COLORS.includes(color) ? '3px solid #323130' : '2px solid #EDEBE9',
                outline: !PROJECT_COLORS.includes(color) ? `2px solid ${color}` : 'none',
                outlineOffset: 2, boxSizing: 'border-box',
              }} />
              <input
                type="color"
                value={color}
                onChange={e => setColor(e.target.value)}
                style={{ position: 'absolute', opacity: 0, width: 0, height: 0, pointerEvents: 'none' }}
              />
            </label>
          </div>
        </div>

        {/* Status */}
        <Dropdown
          label="Status"
          selectedKey={status}
          options={statusOptions}
          onChange={(_, opt) => opt && setStatus(opt.key as ProjectStatus)}
        />

        {/* Date range */}
        <Stack horizontal tokens={{ childrenGap: 12 }}>
          <Stack.Item grow>
            <TextField
              label="Start Date"
              type="date"
              value={startDate}
              onChange={(_, v) => setStartDate(v || '')}
            />
          </Stack.Item>
          <Stack.Item grow>
            <TextField
              label="Due Date"
              type="date"
              value={dueDate}
              onChange={(_, v) => { setDueDate(v || ''); setErrors(p => ({ ...p, dueDate: '' })); }}
              errorMessage={errors.dueDate}
            />
          </Stack.Item>
        </Stack>

        {/* Preview */}
        {title && (
          <div
            style={{
              background: `${color}18`,
              border: `1px solid ${color}40`,
              borderLeft: `4px solid ${color}`,
              borderRadius: 6,
              padding: '10px 14px',
            }}
          >
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <div style={{ width: 12, height: 12, borderRadius: '50%', background: color }} />
              <span style={{ fontWeight: 600, fontSize: 14, color: '#323130' }}>{title}</span>
              <span style={{ fontSize: 11, color: color, fontWeight: 600, marginLeft: 'auto',
                background: `${color}20`, padding: '2px 8px', borderRadius: 10 }}>
                {status}
              </span>
            </div>
            {description && (
              <p style={{ margin: '6px 0 0', fontSize: 12, color: '#605E5C', lineHeight: 1.5 }}>
                {description}
              </p>
            )}
          </div>
        )}

        {!isEdit && (
          <div style={{
            background: '#F3F2F1', borderRadius: 4, padding: '10px 12px',
            fontSize: 12, color: '#605E5C', lineHeight: 1.6,
          }}>
            <strong>What happens next:</strong> A new SharePoint list will be created to store tasks
            for this project. Standard columns (status, priority, dates, assignee, % complete, etc.)
            are added automatically.
          </div>
        )}
      </Stack>
    </Panel>
  );
};

export default ProjectPanel;
