import * as React from 'react';
import {
  Panel, PanelType, PrimaryButton, DefaultButton, Spinner, SpinnerSize, Stack,
  TextField, Dropdown, IDropdownOption, Label,
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
  IImportSource, ColumnMapping, IPlannerPlan,
  parseExcelFile, fetchPlannerPlans, fetchPlannerTasks,
  applyMapping, batchImport, resolveDependencies, IBatchImportResult,
} from '../../services/ImportService';
import { IProject, PROJECT_COLORS, PROJECT_STATUS_OPTIONS, ProjectStatus } from '../../models';
import { SharePointService } from '../../services/SharePointService';
import { ColumnMapper } from './ColumnMapper';
import styles from './ImportPanel.module.scss';

type ImportStep = 'source' | 'project-details' | 'map' | 'importing' | 'done';
type SourceType = 'excel' | 'planner' | null;

interface IImportPanelProps {
  isOpen: boolean;
  /** Omit to create a new project from the imported file. */
  project?: IProject;
  spService: SharePointService;
  context: WebPartContext;
  onDismiss: () => void;
  /** In create-project mode the newly created project is passed back so the
   *  caller can navigate to it. In regular mode the argument is undefined. */
  onImportComplete: (newProject?: IProject) => void;
}

export const ImportPanel: React.FC<IImportPanelProps> = ({
  isOpen, project, spService, context, onDismiss, onImportComplete,
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

  // Import progress
  const [importProgress, setImportProgress] = React.useState({ done: 0, total: 0 });
  const [importResult, setImportResult] = React.useState<IBatchImportResult | null>(null);

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
          setPlansError(e.message || 'Could not load Planner plans. Check that Graph permissions have been approved.');
          setPlansLoading(false);
        });
    }
  }, [sourceType]);

  // ─── Handlers ─────────────────────────────────────────────────────────────

  const handleFileDrop = async (file: File): Promise<void> => {
    setFileError('');
    try {
      const source = await parseExcelFile(file);
      setImportSource(source);
      setMapping(source.autoMapping);

      if (createMode) {
        // Pre-fill project name from filename (strip extension and separators)
        const baseName = file.name
          .replace(/\.(xlsx?|csv|ods)$/i, '')
          .replace(/[-_]+/g, ' ')
          .trim();
        setNewProjectTitle(baseName);
      }
    } catch (e: any) {
      setFileError(e.message || 'Could not parse the file.');
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
    setSelectedPlan(plan);
    setPlanTasksLoading(true);
    try {
      const source = await fetchPlannerTasks(context, plan.id, plan.title);
      setImportSource(source);
      setMapping(source.autoMapping);
      if (createMode) setNewProjectTitle(plan.title);
    } catch (e: any) {
      setPlansError(e.message || 'Could not load tasks from this plan.');
    } finally {
      setPlanTasksLoading(false);
    }
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
    if (!newProjectTitle.trim()) errs.title = 'Project name is required.';
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
    const tasks = applyMapping(importSource.rows, mapping);
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
          errors: [`Could not create project: ${(e as Error).message || 'unknown error'}`],
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

    const result = await batchImport(
      spService,
      targetProject.listName,
      tasks,
      (done, total) => setImportProgress({ done: done + offset, total: total + offset })
    );
    setImportProgress({ done: totalSlots, total: totalSlots });

    // Resolve name-based dependency references (e.g. "Task Title A, Task Title B")
    // to SharePoint numeric IDs now that all tasks have been created.
    if (importSource && Object.values(mapping).includes('dependencies')) {
      await resolveDependencies(spService, targetProject.listName, importSource.rows, mapping);
    }

    setImportResult(result);
    setStep('done');
  };

  const hasTitleMapped = Object.values(mapping).includes('title');
  const taskCount = importSource
    ? importSource.rows.filter(r => {
        const titleCol = Object.keys(mapping).find(k => mapping[k] === 'title');
        return titleCol ? !!r[titleCol]?.trim() : false;
      }).length
    : 0;

  // ─── Step renders ──────────────────────────────────────────────────────────

  const renderSourceStep = (): React.ReactNode => (
    <div>
      <div className={styles.sourceGrid}>
        <div
          className={`${styles.sourceCard} ${sourceType === 'excel' ? styles.selected : ''}`}
          onClick={() => { setSourceType('excel'); setImportSource(null); }}
        >
          <div className={styles.sourceIcon}>📊</div>
          <div className={styles.sourceTitle}>Excel / CSV</div>
          <div className={styles.sourceSubtitle}>
            Upload .xlsx, .xls, or .csv — including exports from MS Project Desktop
          </div>
        </div>
        <div
          className={`${styles.sourceCard} ${sourceType === 'planner' ? styles.selected : ''}`}
          onClick={() => setSourceType('planner')}
        >
          <div className={styles.sourceIcon}>📋</div>
          <div className={styles.sourceTitle}>Microsoft Planner</div>
          <div className={styles.sourceSubtitle}>
            Import tasks directly from any Planner plan in your Microsoft 365 account
          </div>
        </div>
      </div>

      {/* Excel file drop */}
      {sourceType === 'excel' && (
        <>
          <div
            className={`${styles.dropZone} ${dragOver ? styles.dragOver : ''} ${importSource ? styles.hasFile : ''}`}
            onClick={() => fileInputRef.current?.click()}
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
                  {importSource.rows.length} rows · {importSource.headers.length} columns — click to change
                </div>
              </>
            ) : (
              <>
                <div className={styles.dropText}>
                  <strong>Click to browse</strong> or drag &amp; drop your file here
                </div>
                <div className={styles.dropSubtext}>.xlsx · .xls · .csv · .ods</div>
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
            <div style={{ color: '#D13438', fontSize: 13, marginTop: 8 }}>⚠ {fileError}</div>
          )}
        </>
      )}

      {/* Planner plan list */}
      {sourceType === 'planner' && (
        <>
          {plansLoading && (
            <div className={styles.loadingRow}>
              <Spinner size={SpinnerSize.small} />
              Loading your Planner plans…
            </div>
          )}
          {plansError && (
            <div style={{ color: '#D13438', fontSize: 13, padding: '12px 0' }}>
              ⚠ {plansError}
              <div style={{ fontSize: 12, color: '#605E5C', marginTop: 6 }}>
                A Microsoft 365 admin may need to approve <em>Tasks.Read</em> and <em>Group.Read.All</em> permissions for this web part in the SharePoint Admin Center under API Access.
              </div>
            </div>
          )}
          {!plansLoading && !plansError && plans.length === 0 && (
            <div className={styles.emptyPlanner}>
              No Planner plans found. Make sure you belong to at least one Microsoft 365 group that has Planner.
            </div>
          )}
          {!plansLoading && plans.length > 0 && (
            <div className={styles.planList}>
              {plans.map(plan => (
                <div
                  key={plan.id}
                  className={`${styles.planItem} ${selectedPlan?.id === plan.id ? styles.selected : ''}`}
                  onClick={() => void handlePlanSelect(plan)}
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
            task{taskCount !== 1 ? 's' : ''} found in{' '}
            {sourceType === 'excel' ? importSource.fileName : `"${importSource.planName}"`}
          </span>
        </div>
      )}
    </div>
  );

  const statusOptions: IDropdownOption[] = PROJECT_STATUS_OPTIONS.map(s => ({ key: s, text: s }));

  const renderProjectDetailsStep = (): React.ReactNode => (
    <div>
      <div style={{ fontSize: 13, color: '#605E5C', marginBottom: 16 }}>
        A new project will be created and all {taskCount} task{taskCount !== 1 ? 's' : ''} will be imported into it.
        You can change these details any time after import.
      </div>

      <TextField
        label="Project name"
        required
        value={newProjectTitle}
        onChange={(_, v) => setNewProjectTitle(v ?? '')}
        errorMessage={projectErrors.title}
        styles={{ root: { marginBottom: 14 } }}
      />

      <TextField
        label="Description"
        value={newProjectDescription}
        onChange={(_, v) => setNewProjectDescription(v ?? '')}
        multiline
        rows={2}
        styles={{ root: { marginBottom: 14 } }}
      />

      <Dropdown
        label="Status"
        selectedKey={newProjectStatus}
        options={statusOptions}
        onChange={(_, o) => { if (o) setNewProjectStatus(o.key as ProjectStatus); }}
        styles={{ root: { marginBottom: 14 } }}
      />

      <Label>Color</Label>
      <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 4 }}>
        {PROJECT_COLORS.map(c => (
          <div
            key={c}
            onClick={() => setNewProjectColor(c)}
            style={{
              width: 28,
              height: 28,
              borderRadius: '50%',
              background: c,
              cursor: 'pointer',
              border: newProjectColor === c ? '3px solid #323130' : '3px solid transparent',
              outline: newProjectColor === c ? `2px solid ${c}` : 'none',
              outlineOffset: 1,
              transition: 'border 0.1s',
            }}
          />
        ))}
      </div>
    </div>
  );

  const renderMapStep = (): React.ReactNode => (
    <div>
      <div style={{ fontSize: 13, color: '#605E5C', marginBottom: 14 }}>
        Map the columns from your source to Smart Gantt fields. Columns that matched automatically
        are highlighted in green &mdash; adjust any that don&apos;t look right.
      </div>
      {importSource && (
        <ColumnMapper
          source={importSource}
          mapping={mapping}
          onChange={setMapping}
        />
      )}
    </div>
  );

  const renderImportingStep = (): React.ReactNode => {
    const { done, total } = importProgress;
    const pct = total > 0 ? Math.round((done / total) * 100) : 0;
    return (
      <div className={styles.progressSection}>
        <div className={styles.progressTitle}>
          {createMode && done === 0 ? 'Creating project…' : 'Importing tasks…'}
        </div>
        <div className={styles.progressBar} style={{ width: '100%' }}>
          <div className={styles.progressFill} style={{ width: `${pct}%` }} />
        </div>
        <div className={styles.progressLabel}>{done} of {total} ({pct}%)</div>
        <Spinner size={SpinnerSize.medium} />
      </div>
    );
  };

  const renderDoneStep = (): React.ReactNode => {
    if (!importResult) return null;
    const hasErrors = importResult.failed > 0;
    const targetProject = createdProject ?? project;
    return (
      <div className={styles.resultSection}>
        <div className={`${styles.resultCard} ${hasErrors ? styles.partial : styles.success}`}>
          <div className={styles.resultIcon}>{hasErrors ? '⚠️' : '🎉'}</div>
          <div className={styles.resultInfo}>
            <div className={styles.resultTitle}>
              {hasErrors
                ? `Import completed with ${importResult.failed} error${importResult.failed !== 1 ? 's' : ''}`
                : createMode ? 'Project created successfully!' : 'Import successful!'}
            </div>
            <div className={styles.resultDetail}>
              {importResult.succeeded} task{importResult.succeeded !== 1 ? 's' : ''} added to{' '}
              <strong>{targetProject?.title}</strong>
              {hasErrors && ` · ${importResult.failed} failed`}
            </div>
          </div>
        </div>

        {hasErrors && importResult.errors.length > 0 && (
          <div>
            <div style={{ fontSize: 12, fontWeight: 600, color: '#D13438', marginBottom: 6 }}>
              Failed rows:
            </div>
            <div className={styles.errorList}>
              {importResult.errors.map((e, i) => (
                <div key={i} className={styles.errorItem}>• {e}</div>
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
      return (
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton
            text={createMode ? 'Open Project' : 'View Imported Tasks'}
            onClick={() => { onImportComplete(createdProject ?? undefined); onDismiss(); }}
          />
          <DefaultButton text="Close" onClick={onDismiss} />
        </Stack>
      );
    }

    if (step === 'project-details') {
      return (
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton
            text={importSource?.needsMapping || sourceType === 'excel' ? 'Next: Map Columns →' : `Import ${taskCount} Task${taskCount !== 1 ? 's' : ''}`}
            onClick={handleProjectDetailsNext}
          />
          <DefaultButton text="Back" onClick={() => setStep('source')} />
          <DefaultButton text="Cancel" onClick={onDismiss} />
        </Stack>
      );
    }

    if (step === 'map') {
      return (
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton
            text={`Import ${taskCount} Task${taskCount !== 1 ? 's' : ''}`}
            disabled={!hasTitleMapped || taskCount === 0}
            onClick={handleMappingNext}
          />
          <DefaultButton text="Back" onClick={() => setStep(createMode ? 'project-details' : 'source')} />
          <DefaultButton text="Cancel" onClick={onDismiss} />
        </Stack>
      );
    }

    // source step
    const needsMap = importSource?.needsMapping || sourceType === 'excel';
    return (
      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <PrimaryButton
          text={needsMap ? 'Next →' : `Import ${taskCount} Tasks`}
          disabled={!canProceedFromSource() || taskCount === 0}
          onClick={handleNext}
        />
        <DefaultButton text="Cancel" onClick={onDismiss} />
      </Stack>
    );
  };

  // ─── Step indicator ────────────────────────────────────────────────────────

  const steps: Array<{ id: ImportStep; label: string }> = createMode
    ? [
        { id: 'source',          label: 'Source' },
        { id: 'project-details', label: 'Project' },
        { id: 'map',             label: 'Map' },
        { id: 'importing',       label: 'Import' },
        { id: 'done',            label: 'Done' },
      ]
    : [
        { id: 'source',    label: 'Source' },
        { id: 'map',       label: 'Map' },
        { id: 'importing', label: 'Import' },
        { id: 'done',      label: 'Done' },
      ];

  const stepOrder: ImportStep[] = steps.map(s => s.id);
  const currentIdx = stepOrder.indexOf(step);

  const headerText = createMode
    ? 'Import File as New Project'
    : `Import Tasks into "${project!.title}"`;

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
