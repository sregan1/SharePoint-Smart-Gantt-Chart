import * as React from 'react';
import {
  Panel, PanelType, PrimaryButton, DefaultButton, Spinner, SpinnerSize, Stack,
  TextField, Dropdown, IDropdownOption, Label, Checkbox,
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
  IImportSource, ColumnMapping, IPlannerPlan,
  parseExcelFile, fetchPlannerPlans, fetchPlannerTasks,
  applyMapping, filterMappedRows, batchImport, resolveDependencies, IBatchImportResult,
  DateOrder,
} from '../../services/ImportService';
import { IProject, ITask, PROJECT_COLORS, PROJECT_STATUS_OPTIONS, ProjectStatus } from '../../models';
import { SharePointService } from '../../services/SharePointService';
import { ColumnMapper } from './ColumnMapper';
import { ColorSwatchPicker } from '../common/ColorSwatchPicker';
import { onActivate } from '../common/a11y';
import styles from './ImportPanel.module.scss';
import * as strings from 'SmartGanttWebPartStrings';
import { formatString } from '../localeUtils';

type ImportStep = 'source' | 'project-details' | 'map' | 'importing' | 'done';
type SourceType = 'excel' | 'planner' | null;

interface IRowProblem {
  /** Spreadsheet row number (the header is row 1). */
  rowNum: number;
  message: string;
}

// How many row problems the map step lists before summarizing the rest.
const MAX_PROBLEMS_SHOWN = 5;

interface IImportPanelProps {
  isOpen: boolean;
  /** Omit to create a new project from the imported file. */
  project?: IProject;
  /** Existing tasks in `project` (append mode only) — used to keep imported rows'
   *  sort order after the current tasks instead of colliding with them. */
  existingTasks?: ITask[];
  spService: SharePointService;
  context: WebPartContext;
  onDismiss: () => void;
  /** In create-project mode the newly created project is passed back so the
   *  caller can navigate to it. In regular mode the argument is undefined. */
  onImportComplete: (newProject?: IProject) => void;
}

/** Copies text to the clipboard, falling back to a hidden textarea where the
 *  async Clipboard API isn't available (older browsers, non-secure frames). */
async function copyText(text: string): Promise<boolean> {
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch { /* fall through to the legacy path */ }
  try {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    const ok = document.execCommand('copy');
    document.body.removeChild(ta);
    return ok;
  } catch {
    return false;
  }
}

/** Header row for an error/warning list: label on the left, "Copy details" on the right. */
const CopyableListHeader: React.FC<{ label: string; color: string; lines: string[] }> = ({ label, color, lines }) => {
  const [copied, setCopied] = React.useState(false);
  const timer = React.useRef<number | undefined>(undefined);
  React.useEffect(() => () => window.clearTimeout(timer.current), []);
  const onCopy = async (): Promise<void> => {
    if (await copyText(lines.join('\n'))) {
      setCopied(true);
      window.clearTimeout(timer.current);
      timer.current = window.setTimeout(() => setCopied(false), 2000);
    }
  };
  return (
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 6 }}>
      <div style={{ fontSize: 12, fontWeight: 600, color }}>{label}</div>
      <DefaultButton
        text={copied ? strings.Import_Panel_Copied : strings.Import_Panel_CopyDetails}
        iconProps={{ iconName: copied ? 'CheckMark' : 'Copy' }}
        onClick={onCopy}
        styles={{ root: { height: 24, minWidth: 0, padding: '0 8px', fontSize: 12 } }}
        aria-live="polite"
      />
    </div>
  );
};

export const ImportPanel: React.FC<IImportPanelProps> = ({
  isOpen, project, existingTasks, spService, context, onDismiss, onImportComplete,
}) => {
  const createMode = !project;

  const [step, setStep] = React.useState<ImportStep>('source');
  const [sourceType, setSourceType] = React.useState<SourceType>(null);

  // Excel state
  const [dragOver, setDragOver] = React.useState(false);
  const [importSource, setImportSource] = React.useState<IImportSource | null>(null);
  const [mapping, setMapping] = React.useState<ColumnMapping>({});
  const [fileError, setFileError] = React.useState('');
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  // Planner state
  const [plans, setPlans] = React.useState<IPlannerPlan[]>([]);
  const [plansLoading, setPlansLoading] = React.useState(false);
  const [plansError, setPlansError] = React.useState('');
  const [selectedPlan, setSelectedPlan] = React.useState<IPlannerPlan | null>(null);
  const [planTasksLoading, setPlanTasksLoading] = React.useState(false);
  const latestPlanIdRef = React.useRef<string | null>(null);

  // Import progress
  const [importProgress, setImportProgress] = React.useState({ done: 0, total: 0 });
  const [importResult, setImportResult] = React.useState<IBatchImportResult | null>(null);
  // Non-fatal issues (e.g. unresolved dependency links) — kept apart from failed rows
  // so they are still shown when every row imported.
  const [importWarnings, setImportWarnings] = React.useState<string[]>([]);

  // Excel mapping options
  const [dateOrder, setDateOrder] = React.useState<DateOrder>('auto');
  const [skipInvalid, setSkipInvalid] = React.useState(false);

  // New-project state (create mode only)
  const [newProjectTitle, setNewProjectTitle] = React.useState('');
  const [newProjectColor, setNewProjectColor] = React.useState(PROJECT_COLORS[0]);
  const [newProjectStatus, setNewProjectStatus] = React.useState<ProjectStatus>('Active');
  const [newProjectDescription, setNewProjectDescription] = React.useState('');
  const [newProjectStart, setNewProjectStart] = React.useState('');
  const [newProjectEnd, setNewProjectEnd] = React.useState('');
  const [projectErrors, setProjectErrors] = React.useState<Record<string, string>>({});
  const [createdProject, setCreatedProject] = React.useState<IProject | null>(null);

  // Reset on open
  React.useEffect(() => {
    if (isOpen) {
      setStep('source');
      setSourceType(null);
      setImportSource(null);
      setMapping({});
      setFileError('');
      setSelectedPlan(null);
      setPlans([]);
      setPlansError('');
      setImportResult(null);
      setImportWarnings([]);
      setDateOrder('auto');
      setSkipInvalid(false);
      setImportProgress({ done: 0, total: 0 });
      setCreatedProject(null);
      setNewProjectTitle('');
      setNewProjectColor(PROJECT_COLORS[0]);
      setNewProjectStatus('Active');
      setNewProjectDescription('');
      setNewProjectStart('');
      setNewProjectEnd('');
      setProjectErrors({});
    }
  }, [isOpen]);

  // Load Planner plans when user selects Planner source
  React.useEffect(() => {
    if (sourceType === 'planner' && plans.length === 0 && !plansLoading) {
      setPlansLoading(true);
      setPlansError('');
      fetchPlannerPlans(context)
        .then(p => { setPlans(p); setPlansLoading(false); })
        .catch((e: Error) => {
          setPlansError(e.message || strings.ImportPanel_PlannerLoadError);
          setPlansLoading(false);
        });
    }
  }, [sourceType]);

  // Row-level validation, run before importing: missing title, unparseable
  // dates, due before start. Fully blank rows are ignored. Blank-title rows get
  // a placeholder title so applyMapping() keeps the results index-aligned with
  // the source rows.
  const validation = React.useMemo(() => {
    const problems: IRowProblem[] = [];
    const badRows = new Set<number>();
    if (!importSource) return { problems, badRows };
    const titleCol = Object.keys(mapping).find(k => mapping[k] === 'title');
    if (!titleCol) return { problems, badRows };
    const startCol = Object.keys(mapping).find(k => mapping[k] === 'startDate');
    const dueCol = Object.keys(mapping).find(k => mapping[k] === 'dueDate');
    const rows = importSource.rows;
    const aligned = rows.map(r => ((r[titleCol] ?? '').trim() ? r : { ...r, [titleCol]: '\u0000' }));
    const mapped = applyMapping(aligned, mapping, { dateOrder });
    rows.forEach((r, i) => {
      const hasContent = Object.keys(r).some(k => String(r[k] ?? '').trim() !== '');
      if (!hasContent) return;
      const t = mapped[i];
      const messages: string[] = [];
      if (!(r[titleCol] ?? '').trim()) messages.push(strings.Import_Panel_ProblemMissingTitle);
      if (startCol && (r[startCol] ?? '').trim() && !t?.startDate) {
        messages.push(formatString(strings.Import_Panel_ProblemInvalidStartDate, { value: r[startCol].trim() }));
      }
      if (dueCol && (r[dueCol] ?? '').trim() && !t?.dueDate) {
        messages.push(formatString(strings.Import_Panel_ProblemInvalidDueDate, { value: r[dueCol].trim() }));
      }
      if (t?.startDate && t?.dueDate && t.dueDate < t.startDate) messages.push(strings.Import_Panel_ProblemEndBeforeStart);
      if (messages.length > 0) {
        badRows.add(i);
        problems.push({ rowNum: i + 2, message: messages.join(' · ') });
      }
    });
    return { problems, badRows };
  }, [importSource, mapping, dateOrder]);

  // ─── Handlers ─────────────────────────────────────────────────────────────

  const handleFileDrop = async (file: File): Promise<void> => {
    setFileError('');
    try {
      const source = await parseExcelFile(file);
      setImportSource(source);
      setMapping(source.autoMapping);
      setDateOrder('auto');
      setSkipInvalid(false);

      if (createMode) {
        // Pre-fill project name from filename (strip extension and separators)
        const baseName = file.name
          .replace(/\.(xlsx?|csv|ods)$/i, '')
          .replace(/[-_]+/g, ' ')
          .trim();
        setNewProjectTitle(baseName);
      }
    } catch (e: any) {
      setFileError(e.message || strings.ImportPanel_CouldNotParseFile);
    }
  };

  const handleDropZoneDrop = (e: React.DragEvent): void => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files[0];
    if (file) void handleFileDrop(file);
  };

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files?.[0];
    if (file) void handleFileDrop(file);
    e.target.value = '';
  };

  const handlePlanSelect = async (plan: IPlannerPlan): Promise<void> => {
    latestPlanIdRef.current = plan.id;
    setSelectedPlan(plan);
    setPlanTasksLoading(true);
    try {
      const source = await fetchPlannerTasks(context, plan.id, plan.title);
      // A later click on a different plan may have resolved first; don't let
      // this (now-stale) response overwrite it.
      if (latestPlanIdRef.current !== plan.id) return;
      setImportSource(source);
      setMapping(source.autoMapping);
      if (createMode) setNewProjectTitle(plan.title);
    } catch (e: any) {
      if (latestPlanIdRef.current !== plan.id) return;
      setPlansError(e.message || strings.ImportPanel_CouldNotLoadPlannerTasks);
    } finally {
      if (latestPlanIdRef.current === plan.id) setPlanTasksLoading(false);
    }
  };

  const selectSource = (type: 'excel' | 'planner'): void => {
    setSourceType(type);
    setImportSource(null);
    setMapping({});
    setSelectedPlan(null);
    latestPlanIdRef.current = null;
  };

  const canProceedFromSource = (): boolean => {
    if (sourceType === 'excel') return !!importSource;
    if (sourceType === 'planner') return !!importSource;
    return false;
  };

  const handleNext = (): void => {
    if (!importSource) return;
    if (createMode) {
      setStep('project-details');
    } else if (importSource.needsMapping || sourceType === 'excel') {
      setStep('map');
    } else {
      void startImport();
    }
  };

  const handleProjectDetailsNext = (): void => {
    const errs: Record<string, string> = {};
    if (!newProjectTitle.trim()) errs.title = strings.ProjectPanel_ProjectNameRequired;
    if (newProjectStart && newProjectEnd && newProjectEnd < newProjectStart) errs.end = strings.ProjectPanel_DueDateError;
    setProjectErrors(errs);
    if (Object.keys(errs).length > 0) return;

    if (importSource?.needsMapping || sourceType === 'excel') {
      setStep('map');
    } else {
      void startImport();
    }
  };

  const handleMappingNext = (): void => {
    void startImport();
  };

  const startImport = async (): Promise<void> => {
    if (!importSource) return;
    // Keep the raw rows used for dependency resolution in lockstep with the
    // task array batchImport() creates from, so createdIds line up positionally.
    // Skipped rows keep their place (blank title) instead of being removed, so
    // spreadsheet row numbers used by MS Project-style dependencies stay valid.
    const titleColumn = Object.keys(mapping).find(k => mapping[k] === 'title');
    const sourceRows = skipInvalid && titleColumn
      ? importSource.rows.map((r, i) => (validation.badRows.has(i) ? { ...r, [titleColumn]: '' } : r))
      : importSource.rows;
    const filteredRows = filterMappedRows(sourceRows, mapping);
    const tasks = applyMapping(sourceRows, mapping, { dateOrder });
    if (tasks.length === 0) return;

    setStep('importing');

    let targetProject = project ?? null;

    if (createMode) {
      // Step 1: create the project
      const extraSlots = 1;
      setImportProgress({ done: 0, total: tasks.length + extraSlots });
      try {
        targetProject = await spService.createProject({
          title: newProjectTitle.trim(),
          description: newProjectDescription.trim(),
          color: newProjectColor,
          startDate: newProjectStart || '',
          dueDate: newProjectEnd || '',
          status: newProjectStatus,
        });
        setCreatedProject(targetProject);
        setImportProgress({ done: 1, total: tasks.length + extraSlots });
      } catch (e: any) {
        setImportResult({
          succeeded: 0,
          failed: tasks.length,
          errors: [formatString(strings.ImportPanel_CouldNotCreateProjectDetail, { detail: (e as Error).message || strings.ImportPanel_UnknownError })],
          createdIds: [],
        });
        setStep('done');
        return;
      }
    } else {
      setImportProgress({ done: 0, total: tasks.length });
    }

    if (!targetProject) return;

    const offset = createMode ? 1 : 0;
    const totalSlots = tasks.length + offset;
    // Append after any tasks already in the project, so imported rows don't
    // interleave with the existing manual sort order.
    const sortOrderBase = (createMode || !existingTasks || existingTasks.length === 0)
      ? 0
      : Math.max(...existingTasks.map(t => t.sortOrder)) + 1;

    const result = await batchImport(
      spService,
      targetProject.listName,
      tasks,
      sortOrderBase,
      (done, total) => setImportProgress({ done: done + offset, total: total + offset })
    );
    setImportProgress({ done: totalSlots, total: totalSlots });

    // Resolve name-based / row-number dependency references now that all
    // tasks have been created. Tasks are already imported at this point, so
    // a failure here is reported as a warning rather than stranding the
    // wizard on the importing step.
    const warnings: string[] = [];
    if (Object.values(mapping).includes('dependencies')) {
      try {
        const depResult = await resolveDependencies(
          spService, targetProject.listName, filteredRows, mapping, result.createdIds
        );
        if (depResult.warnings.length > 0) warnings.push(...depResult.warnings);
      } catch (e) {
        warnings.push(formatString(strings.ImportPanel_CouldNotLinkDependencies, { detail: e instanceof Error ? e.message : strings.ImportPanel_UnknownError }));
      }
    }

    setImportWarnings(warnings);
    setImportResult(result);
    setStep('done');
  };

  const hasTitleMapped = Object.values(mapping).includes('title');
  const taskCount = importSource
    ? importSource.rows.filter((r, i) => {
        const titleCol = Object.keys(mapping).find(k => mapping[k] === 'title');
        if (!titleCol || !r[titleCol]?.trim()) return false;
        return !(skipInvalid && validation.badRows.has(i));
      }).length
    : 0;
  const hasDateColumns = Object.values(mapping).some(v => v === 'startDate' || v === 'dueDate');

  // ─── Step renders ──────────────────────────────────────────────────────────

  const renderSourceStep = (): React.ReactNode => (
    <div>
      <div className={styles.sourceGrid}>
        <div
          className={`${styles.sourceCard} ${sourceType === 'excel' ? styles.selected : ''}`}
          onClick={() => selectSource('excel')}
          role="button"
          tabIndex={0}
          aria-pressed={sourceType === 'excel'}
          onKeyDown={onActivate(() => selectSource('excel'))}
        >
          <div className={styles.sourceIcon}>📊</div>
          <div className={styles.sourceTitle}>{strings.ImportPanel_SourceExcelTitle}</div>
          <div className={styles.sourceSubtitle}>
            {strings.ImportPanel_SourceExcelSubtitle}
          </div>
        </div>
        <div
          className={`${styles.sourceCard} ${sourceType === 'planner' ? styles.selected : ''}`}
          onClick={() => selectSource('planner')}
          role="button"
          tabIndex={0}
          aria-pressed={sourceType === 'planner'}
          onKeyDown={onActivate(() => selectSource('planner'))}
        >
          <div className={styles.sourceIcon}>📋</div>
          <div className={styles.sourceTitle}>{strings.ImportPanel_SourcePlannerTitle}</div>
          <div className={styles.sourceSubtitle}>
            {strings.ImportPanel_SourcePlannerSubtitle}
          </div>
        </div>
      </div>

      {/* Excel file drop */}
      {sourceType === 'excel' && (
        <>
          <div
            className={`${styles.dropZone} ${dragOver ? styles.dragOver : ''} ${importSource ? styles.hasFile : ''}`}
            onClick={() => fileInputRef.current?.click()}
            role="button"
            tabIndex={0}
            aria-label={importSource ? formatString(strings.Import_Panel_DropZoneChangeAriaLabel, { fileName: importSource.fileName || '' }) : strings.Import_Panel_DropZoneAriaLabel}
            onKeyDown={onActivate(() => fileInputRef.current?.click())}
            onDragOver={e => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={handleDropZoneDrop}
          >
            <div className={styles.dropIcon}>
              {importSource ? '✅' : '📁'}
            </div>
            {importSource ? (
              <>
                <div className={styles.dropText}>
                  <strong>{importSource.fileName}</strong>
                </div>
                <div className={styles.dropSubtext}>
                  {formatString(strings.ImportPanel_RowsColsSummary, { rows: importSource.rows.length, cols: importSource.headers.length })}
                </div>
              </>
            ) : (
              <>
                <div className={styles.dropText}>
                  <strong>{strings.ImportPanel_ClickToBrowse}</strong>{strings.ImportPanel_DragDropSuffix}
                </div>
                <div className={styles.dropSubtext}>{strings.ImportPanel_SupportedExtensions}</div>
              </>
            )}
          </div>
          <input
            ref={fileInputRef}
            type="file"
            className={styles.fileInput}
            accept=".xlsx,.xls,.csv,.ods"
            onChange={handleFileInputChange}
          />
          {fileError && (
            <div role="alert" style={{ color: '#D13438', fontSize: 13, marginTop: 8 }}>⚠ {fileError}</div>
          )}
        </>
      )}

      {/* Planner plan list */}
      {sourceType === 'planner' && (
        <>
          {plansLoading && (
            <div className={styles.loadingRow}>
              <Spinner size={SpinnerSize.small} />
              {strings.ImportPanel_LoadingPlannerPlans}
            </div>
          )}
          {plansError && (
            <div style={{ color: '#D13438', fontSize: 13, padding: '12px 0' }}>
              ⚠ {plansError}
              <div style={{ fontSize: 12, color: 'var(--neutralSecondary, #605E5C)', marginTop: 6 }}>
                {strings.ImportPanel_PlannerPermissionsHintPrefix}<em>{strings.ImportPanel_PlannerPermissionsHintEmphasis}</em>{strings.ImportPanel_PlannerPermissionsHintSuffix}
              </div>
            </div>
          )}
          {!plansLoading && !plansError && plans.length === 0 && (
            <div className={styles.emptyPlanner}>
              <div>{strings.ImportPanel_NoPlannerPlansFoundTitle}</div>
              <div>{strings.ImportPanel_NoPlannerPlansFoundBody}</div>
            </div>
          )}
          {!plansLoading && plans.length > 0 && (
            <div className={styles.planList}>
              {plans.map(plan => (
                <div
                  key={plan.id}
                  className={`${styles.planItem} ${selectedPlan?.id === plan.id ? styles.selected : ''}`}
                  onClick={() => void handlePlanSelect(plan)}
                  role="button"
                  tabIndex={0}
                  aria-pressed={selectedPlan?.id === plan.id}
                  aria-label={formatString(strings.Import_Panel_PlanAriaLabel, { planName: plan.title, groupName: plan.groupName })}
                  onKeyDown={onActivate(() => void handlePlanSelect(plan))}
                >
                  <div className={styles.planIcon}>📋</div>
                  <div className={styles.planInfo}>
                    <div className={styles.planTitle}>{plan.title}</div>
                    <div className={styles.planGroup}>{plan.groupName}</div>
                  </div>
                  {planTasksLoading && selectedPlan?.id === plan.id && (
                    <Spinner size={SpinnerSize.small} />
                  )}
                  {selectedPlan?.id === plan.id && importSource && (
                    <span className={styles.planCheck}>✓</span>
                  )}
                </div>
              ))}
            </div>
          )}
        </>
      )}

      {/* Task count summary */}
      {importSource && taskCount > 0 && (
        <div className={styles.previewSummary} style={{ marginTop: 12 }}>
          <span className={styles.previewCount}>{taskCount}</span>
          <span className={styles.previewCountLabel}>
            {formatString(strings.ImportPanel_TaskCountFoundIn, {
              count: taskCount,
              planName: sourceType === 'excel' ? importSource.fileName : importSource.planName,
            })}
          </span>
        </div>
      )}
    </div>
  );

  const statusLabels: Record<ProjectStatus, string> = {
    'Planning': strings.ProjectStatus_Planning,
    'Active': strings.ProjectStatus_Active,
    'On Hold': strings.ProjectStatus_OnHold,
    'Completed': strings.ProjectStatus_Completed,
    'Cancelled': strings.ProjectStatus_Cancelled,
  };
  const statusOptions: IDropdownOption[] = PROJECT_STATUS_OPTIONS.map(s => ({ key: s, text: statusLabels[s] }));

  const renderProjectDetailsStep = (): React.ReactNode => (
    <div>
      <div style={{ fontSize: 13, color: 'var(--neutralSecondary, #605E5C)', marginBottom: 16 }}>
        {formatString(taskCount === 1 ? strings.Import_Panel_NewProjectIntroOne : strings.Import_Panel_NewProjectIntroMany, { count: taskCount })}
      </div>

      <TextField
        label={strings.ImportPanel_ProjectNameLabel}
        required
        value={newProjectTitle}
        onChange={(_, v) => setNewProjectTitle(v ?? '')}
        errorMessage={projectErrors.title}
        styles={{ root: { marginBottom: 14 } }}
      />

      <TextField
        label={strings.ImportPanel_DescriptionLabel}
        value={newProjectDescription}
        onChange={(_, v) => setNewProjectDescription(v ?? '')}
        multiline
        rows={2}
        styles={{ root: { marginBottom: 14 } }}
      />

      <Dropdown
        label={strings.ImportPanel_StatusLabel}
        selectedKey={newProjectStatus}
        options={statusOptions}
        onChange={(_, o) => { if (o) setNewProjectStatus(o.key as ProjectStatus); }}
        styles={{ root: { marginBottom: 14 } }}
      />

      <Stack horizontal tokens={{ childrenGap: 12 }} styles={{ root: { marginBottom: 14 } }}>
        <Stack.Item grow>
          <TextField
            label={strings.ProjectPanel_StartDateLabel}
            type="date"
            value={newProjectStart}
            onChange={(_, v) => { setNewProjectStart(v ?? ''); setProjectErrors(prev => ({ ...prev, end: '' })); }}
          />
        </Stack.Item>
        <Stack.Item grow>
          <TextField
            label={strings.ProjectPanel_DueDateLabel}
            type="date"
            value={newProjectEnd}
            onChange={(_, v) => { setNewProjectEnd(v ?? ''); setProjectErrors(prev => ({ ...prev, end: '' })); }}
            errorMessage={projectErrors.end}
          />
        </Stack.Item>
      </Stack>

      <Label>{strings.ImportPanel_ColorLabel}</Label>
      <ColorSwatchPicker
        colors={PROJECT_COLORS}
        value={newProjectColor}
        onChange={setNewProjectColor}
        ariaLabel={strings.ImportPanel_ColorLabel}
      />
    </div>
  );

  const dateOrderOptions: IDropdownOption[] = [
    { key: 'auto', text: strings.Import_Panel_DateOrderAuto },
    { key: 'mdy', text: strings.Import_Panel_DateOrderMDY },
    { key: 'dmy', text: strings.Import_Panel_DateOrderDMY },
  ];

  const renderMapStep = (): React.ReactNode => (
    <div>
      <div style={{ fontSize: 13, color: 'var(--neutralSecondary, #605E5C)', marginBottom: 14 }}>
        {strings.ImportPanel_MapIntro}
      </div>
      {importSource && (
        <ColumnMapper
          source={importSource}
          mapping={mapping}
          onChange={setMapping}
        />
      )}

      {/* Ambiguous numeric dates (03/04/2026) — let the user override detection */}
      {importSource && sourceType === 'excel' && hasDateColumns && (
        <div style={{ marginTop: 16 }}>
          <Dropdown
            label={strings.Import_Panel_DateOrderLabel}
            selectedKey={dateOrder}
            options={dateOrderOptions}
            onChange={(_, o) => { if (o) setDateOrder(o.key as DateOrder); }}
          />
          <div style={{ fontSize: 11, color: 'var(--neutralSecondary, #605E5C)', marginTop: 4 }}>
            {strings.Import_Panel_DateOrderHint}
            {dateOrder === 'auto' && importSource.detectedDateOrder && (
              <> {formatString(strings.Import_Panel_DateOrderDetected, {
                order: importSource.detectedDateOrder === 'dmy' ? strings.Import_Panel_DateOrderDMY : strings.Import_Panel_DateOrderMDY,
              })}</>
            )}
          </div>
        </div>
      )}

      {/* Row-level problems, found before anything is written */}
      {importSource && hasTitleMapped && validation.problems.length > 0 && (
        <div style={{ marginTop: 16 }} role="region" aria-label={strings.Import_Panel_ValidationHeader}>
          <div style={{ fontSize: 12, fontWeight: 600, color: '#CA5010', marginBottom: 6 }}>
            ⚠ {formatString(validation.problems.length === 1 ? strings.Import_Panel_ValidationSummaryOne : strings.Import_Panel_ValidationSummaryMany, { count: validation.problems.length })}
          </div>
          <div className={styles.errorList}>
            {validation.problems.slice(0, MAX_PROBLEMS_SHOWN).map(pr => (
              <div key={pr.rowNum} className={styles.errorItem}>
                • {formatString(strings.Import_Panel_ValidationRow, { rowNum: pr.rowNum, problem: pr.message })}
              </div>
            ))}
            {validation.problems.length > MAX_PROBLEMS_SHOWN && (
              <div className={styles.errorItem}>
                {formatString(strings.Import_Panel_ValidationMore, { count: validation.problems.length - MAX_PROBLEMS_SHOWN })}
              </div>
            )}
          </div>
          <Checkbox
            label={strings.Import_Panel_SkipInvalidRows}
            checked={skipInvalid}
            onChange={(_, checked) => setSkipInvalid(!!checked)}
            styles={{ root: { marginTop: 8 } }}
          />
        </div>
      )}
    </div>
  );

  const renderImportingStep = (): React.ReactNode => {
    const { done, total } = importProgress;
    const pct = total > 0 ? Math.round((done / total) * 100) : 0;
    return (
      <div className={styles.progressSection} role="status" aria-live="polite">
        <div className={styles.progressTitle}>
          {createMode && done === 0 ? strings.ImportPanel_CreatingProjectStatus : strings.ImportPanel_ImportingTasksStatus}
        </div>
        <div
          className={styles.progressBar}
          style={{ width: '100%' }}
          role="progressbar"
          aria-label={strings.ImportPanel_ImportingTasksStatus}
          aria-valuemin={0}
          aria-valuemax={100}
          aria-valuenow={pct}
        >
          <div className={styles.progressFill} style={{ width: `${pct}%` }} />
        </div>
        <div className={styles.progressLabel}>{formatString(strings.Import_Panel_ProgressLabelWithPercent, { done, total, percent: pct })}</div>
        <Spinner size={SpinnerSize.medium} />
      </div>
    );
  };

  const renderDoneStep = (): React.ReactNode => {
    if (!importResult) return null;
    const hasErrors = importResult.failed > 0;
    const hasWarnings = importWarnings.length > 0;
    const targetProject = createdProject ?? project;
    // Project creation itself failed (create mode only) — nothing was
    // imported, so don't claim "0 tasks added to <blank project>".
    const projectCreateFailed = createMode && !createdProject;
    return (
      <div className={styles.resultSection}>
        <div className={`${styles.resultCard} ${hasErrors || hasWarnings ? styles.partial : styles.success}`} role="status">
          <div className={styles.resultIcon}>{projectCreateFailed ? '❌' : hasErrors || hasWarnings ? '⚠️' : '🎉'}</div>
          <div className={styles.resultInfo}>
            <div className={styles.resultTitle}>
              {projectCreateFailed
                ? strings.ImportPanel_CouldNotCreateProject
                : hasErrors
                ? formatString(strings.ImportPanel_ImportCompletedWithErrors, { count: importResult.failed })
                : hasWarnings
                ? formatString(importWarnings.length === 1 ? strings.Import_Panel_SuccessWithWarningsOne : strings.Import_Panel_SuccessWithWarningsMany, { count: importWarnings.length })
                : createMode ? strings.ImportPanel_ProjectCreatedSuccess : strings.ImportPanel_ImportSuccessful}
            </div>
            {!projectCreateFailed && (
              <div className={styles.resultDetail}>
                {formatString(strings.ImportPanel_TasksAddedTo, { count: importResult.succeeded, projectName: '‹PN›' })
                  .split('‹PN›')
                  .map((part, i, arr) => (
                    <React.Fragment key={i}>
                      {part}
                      {i < arr.length - 1 && <strong>{targetProject?.title}</strong>}
                    </React.Fragment>
                  ))}
                {hasErrors && ` · ${formatString(strings.Import_Panel_FailedCount, { count: importResult.failed })}`}
              </div>
            )}
          </div>
        </div>

        {importResult.errors.length > 0 && (
          <div>
            <CopyableListHeader label={strings.ImportPanel_FailedRowsLabel} color="#D13438" lines={importResult.errors} />
            <div className={styles.errorList}>
              {importResult.errors.map((e, i) => (
                <div key={i} className={styles.errorItem}>• {e}</div>
              ))}
            </div>
          </div>
        )}

        {hasWarnings && (
          <div>
            <CopyableListHeader label={formatString(strings.Import_Panel_WarningsLabel, { count: importWarnings.length })} color="#CA5010" lines={importWarnings} />
            <div className={styles.errorList}>
              {importWarnings.map((w, i) => (
                <div key={i} className={styles.errorItem}>• {w}</div>
              ))}
            </div>
          </div>
        )}
      </div>
    );
  };

  // ─── Footer ────────────────────────────────────────────────────────────────

  const renderFooter = (): JSX.Element => {
    if (step === 'importing') return <></>;

    if (step === 'done') {
      if (createMode && !createdProject) {
        return <DefaultButton text={strings.ImportPanel_CloseButton} onClick={onDismiss} />;
      }
      return (
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton
            text={createMode ? strings.ImportPanel_OpenProjectButton : strings.ImportPanel_ViewImportedTasksButton}
            onClick={() => { onImportComplete(createdProject ?? undefined); onDismiss(); }}
          />
          <DefaultButton text={strings.ImportPanel_CloseButton} onClick={onDismiss} />
        </Stack>
      );
    }

    if (step === 'project-details') {
      return (
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton
            text={importSource?.needsMapping || sourceType === 'excel' ? strings.ImportPanel_NextMapColumnsButton : formatString(strings.ImportPanel_ImportNTasksButton, { count: taskCount })}
            onClick={handleProjectDetailsNext}
          />
          <DefaultButton text={strings.ImportPanel_BackButton} onClick={() => setStep('source')} />
          <DefaultButton text={strings.ImportPanel_CancelButton} onClick={onDismiss} />
        </Stack>
      );
    }

    if (step === 'map') {
      return (
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton
            text={formatString(strings.ImportPanel_ImportNTasksButton, { count: taskCount })}
            disabled={!hasTitleMapped || taskCount === 0}
            onClick={handleMappingNext}
          />
          <DefaultButton text={strings.ImportPanel_BackButton} onClick={() => setStep(createMode ? 'project-details' : 'source')} />
          <DefaultButton text={strings.ImportPanel_CancelButton} onClick={onDismiss} />
        </Stack>
      );
    }

    // source step
    const needsMap = importSource?.needsMapping || sourceType === 'excel';
    return (
      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <PrimaryButton
          text={needsMap ? strings.ImportPanel_NextButton : formatString(strings.ImportPanel_ImportNTasksButton, { count: taskCount })}
          disabled={!canProceedFromSource() || taskCount === 0}
          onClick={handleNext}
        />
        <DefaultButton text={strings.ImportPanel_CancelButton} onClick={onDismiss} />
      </Stack>
    );
  };

  // ─── Step indicator ────────────────────────────────────────────────────────

  const steps: Array<{ id: ImportStep; label: string }> = createMode
    ? [
        { id: 'source',          label: strings.ImportPanel_StepSource },
        { id: 'project-details', label: strings.ImportPanel_StepProject },
        { id: 'map',             label: strings.ImportPanel_StepMap },
        { id: 'importing',       label: strings.ImportPanel_StepImport },
        { id: 'done',            label: strings.ImportPanel_StepDone },
      ]
    : [
        { id: 'source',    label: strings.ImportPanel_StepSource },
        { id: 'map',       label: strings.ImportPanel_StepMap },
        { id: 'importing', label: strings.ImportPanel_StepImport },
        { id: 'done',      label: strings.ImportPanel_StepDone },
      ];

  const stepOrder: ImportStep[] = steps.map(s => s.id);
  const currentIdx = stepOrder.indexOf(step);

  const headerText = createMode
    ? strings.ImportPanel_ImportAsNewProjectHeader
    : formatString(strings.ImportPanel_ImportIntoProjectHeader, { projectName: project!.title });

  return (
    <Panel
      isOpen={isOpen}
      type={PanelType.medium}
      headerText={headerText}
      onDismiss={onDismiss}
      isFooterAtBottom
      onRenderFooterContent={renderFooter}
    >
      <div className={styles.importPanel}>
        {/* Step indicator */}
        <div className={styles.stepBar}>
          {steps.map((s, i) => {
            const idx = stepOrder.indexOf(s.id);
            const state = idx < currentIdx ? 'done' : idx === currentIdx ? 'active' : 'pending';
            const isLast = i === steps.length - 1;

            return (
              <React.Fragment key={s.id}>
                <div className={styles.step}>
                  <div className={`${styles.stepCircle} ${styles[state]}`}>
                    {state === 'done' ? '✓' : i + 1}
                  </div>
                  <span className={`${styles.stepLabel} ${styles[state]}`}>{s.label}</span>
                </div>
                {!isLast && (
                  <div className={`${styles.stepConnector} ${state === 'done' ? styles.done : ''}`} />
                )}
              </React.Fragment>
            );
          })}
        </div>

        {/* Step content */}
        {step === 'source'          && renderSourceStep()}
        {step === 'project-details' && renderProjectDetailsStep()}
        {step === 'map'             && renderMapStep()}
        {step === 'importing'       && renderImportingStep()}
        {step === 'done'            && renderDoneStep()}
      </div>
    </Panel>
  );
};

export default ImportPanel;
