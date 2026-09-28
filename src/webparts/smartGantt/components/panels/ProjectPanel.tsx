import * as React from 'react';
import {
  Panel, PanelType, TextField, Dropdown, IDropdownOption,
  PrimaryButton, DefaultButton, Stack, Label, Spinner, SpinnerSize,
  MessageBar, MessageBarType,
} from '@fluentui/react';
import { IProject, PROJECT_COLORS, PROJECT_STATUS_OPTIONS, ProjectStatus } from '../../models';
import { toDateOnly } from '../../utils/dateUtils';
import { ColorSwatchPicker } from '../common/ColorSwatchPicker';
import * as strings from 'SmartGanttWebPartStrings';

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
  const [saveError, setSaveError] = React.useState('');
  const [dirty, setDirty] = React.useState(false);

  // Populate form when editing
  React.useEffect(() => {
    if (isOpen) {
      if (project) {
        setTitle(project.title);
        setDescription(project.description);
        setColor(project.color || PROJECT_COLORS[0]);
        setStartDate(toDateOnly(project.startDate));
        setDueDate(toDateOnly(project.dueDate));
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
      setSaveError('');
      setSaving(false);
      setDirty(false);
    }
  }, [isOpen, project]);

  const validate = (): boolean => {
    const errs: Record<string, string> = {};
    if (!title.trim()) errs.title = strings.ProjectPanel_ProjectNameRequired;
    if (dueDate && startDate && dueDate < startDate) errs.dueDate = strings.ProjectPanel_DueDateError;
    setErrors(errs);
    return Object.keys(errs).length === 0;
  };

  const handleSave = async (): Promise<void> => {
    if (!validate()) return;
    setSaving(true);
    setSaveError('');
    try {
      await onSave({
        title: title.trim(),
        description,
        color,
        startDate,
        dueDate,
        status,
        etag: project?.etag,
      });
    } catch (e) {
      setSaveError(e instanceof Error ? e.message : strings.ProjectPanel_SaveFallbackError);
    } finally {
      setSaving(false);
    }
  };

  const handleDismiss = (): void => {
    if (dirty && !saving && !window.confirm(strings.ProjectPanel_DiscardUnsavedChangesConfirm)) return;
    onDismiss();
  };

  const statusLabels: Record<ProjectStatus, string> = {
    'Planning': strings.ProjectStatus_Planning,
    'Active': strings.ProjectStatus_Active,
    'On Hold': strings.ProjectStatus_OnHold,
    'Completed': strings.ProjectStatus_Completed,
    'Cancelled': strings.ProjectStatus_Cancelled,
  };
  const statusOptions: IDropdownOption[] = PROJECT_STATUS_OPTIONS.map(s => ({ key: s, text: statusLabels[s] }));

  return (
    <Panel
      isOpen={isOpen}
      type={PanelType.smallFixedFar}
      headerText={isEdit ? strings.ProjectPanel_EditHeader : strings.ProjectPanel_NewHeader}
      onDismiss={handleDismiss}
      isFooterAtBottom
      onRenderFooterContent={() => (
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton
            text={saving ? '' : isEdit ? strings.ProjectPanel_SaveChangesButton : strings.ProjectPanel_CreateProjectButton}
            onClick={handleSave}
            disabled={saving}
          >
            {saving && <Spinner size={SpinnerSize.small} style={{ marginRight: 6 }} />}
            {saving ? (isEdit ? strings.ProjectPanel_SavingLabel : strings.ProjectPanel_CreatingLabel) : undefined}
          </PrimaryButton>
          <DefaultButton text={strings.ProjectPanel_CancelButton} onClick={handleDismiss} disabled={saving} />
        </Stack>
      )}
    >
      {saveError && (
        <MessageBar
          messageBarType={MessageBarType.error}
          onDismiss={() => setSaveError('')}
          dismissButtonAriaLabel={strings.ProjectPanel_DismissAriaLabel}
          styles={{ root: { marginTop: 12 } }}
        >
          {saveError}
        </MessageBar>
      )}
      <Stack tokens={{ childrenGap: 16 }} style={{ marginTop: 20 }}>
        {/* Project name */}
        <TextField
          label={strings.ProjectPanel_ProjectNameLabel}
          value={title}
          onChange={(_, v) => { setTitle(v || ''); setErrors(p => ({ ...p, title: '' })); setDirty(true); }}
          required
          errorMessage={errors.title}
          autoFocus
          placeholder={strings.ProjectPanel_ProjectNamePlaceholder}
        />

        {/* Description */}
        <TextField
          label={strings.ProjectPanel_DescriptionLabel}
          value={description}
          onChange={(_, v) => { setDescription(v || ''); setDirty(true); }}
          multiline
          rows={3}
          placeholder={strings.ProjectPanel_DescriptionPlaceholder}
          resizable={false}
        />

        {/* Color picker */}
        <div>
          <Label>{strings.ProjectPanel_ProjectColorLabel}</Label>
          <ColorSwatchPicker
            colors={PROJECT_COLORS}
            value={color}
            onChange={c => { setColor(c); setDirty(true); }}
            size={30}
          />
        </div>

        {/* Status */}
        <Dropdown
          label={strings.ProjectPanel_StatusLabel}
          selectedKey={status}
          options={statusOptions}
          onChange={(_, opt) => { if (opt) { setStatus(opt.key as ProjectStatus); setDirty(true); } }}
        />

        {/* Date range */}
        <Stack horizontal tokens={{ childrenGap: 12 }}>
          <Stack.Item grow>
            <TextField
              label={strings.ProjectPanel_StartDateLabel}
              type="date"
              value={startDate}
              onChange={(_, v) => { setStartDate(v || ''); setDirty(true); }}
            />
          </Stack.Item>
          <Stack.Item grow>
            <TextField
              label={strings.ProjectPanel_DueDateLabel}
              type="date"
              value={dueDate}
              onChange={(_, v) => { setDueDate(v || ''); setErrors(p => ({ ...p, dueDate: '' })); setDirty(true); }}
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
                {statusLabels[status]}
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
            <strong>{strings.ProjectPanel_WhatHappensNextTitle}</strong> {strings.ProjectPanel_WhatHappensNextBody}
          </div>
        )}
      </Stack>
    </Panel>
  );
};

export default ProjectPanel;
