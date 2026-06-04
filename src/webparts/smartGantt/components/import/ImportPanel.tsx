import * as React from 'react';
import { Panel, PanelType, PrimaryButton, DefaultButton, Spinner, SpinnerSize, Stack } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
  IImportSource, ColumnMapping, IPlannerPlan,
  parseExcelFile, fetchPlannerPlans, fetchPlannerTasks,
  applyMapping, batchImport, IBatchImportResult,
} from '../../services/ImportService';
import { IProject } from '../../models';
import { SharePointService } from '../../services/SharePointService';
import { ColumnMapper } from './ColumnMapper';
import styles from './ImportPanel.module.scss';

type ImportStep = 'source' | 'map' | 'importing' | 'done';
type SourceType = 'excel' | 'planner' | null;

interface IImportPanelProps {
  isOpen: boolean;
  project: IProject;
  spService: SharePointService;
  context: WebPartContext;
  onDismiss: () => void;
  onImportComplete: () => void;
}

export const ImportPanel: React.FC<IImportPanelProps> = ({
  isOpen, project, spService, context, onDismiss, onImportComplete,
}) => {
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
    if (importSource.needsMapping || sourceType === 'excel') {
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

    setImportProgress({ done: 0, total: tasks.length });
    setStep('importing');

    const result = await batchImport(
      spService,
      project.listName,
      tasks,
      (done, total) => setImportProgress({ done, total })
    );

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
        <div className={styles.progressTitle}>Importing tasks…</div>
        <div className={styles.progressBar} style={{ width: '100%' }}>
          <div className={styles.progressFill} style={{ width: `${pct}%` }} />
        </div>
        <div className={styles.progressLabel}>{done} of {total} tasks created ({pct}%)</div>
        <Spinner size={SpinnerSize.medium} />
      </div>
    );
  };

  const renderDoneStep = (): React.ReactNode => {
    if (!importResult) return null;
    const hasErrors = importResult.failed > 0;
    return (
      <div className={styles.resultSection}>
        <div className={`${styles.resultCard} ${hasErrors ? styles.partial : styles.success}`}>
          <div className={styles.resultIcon}>{hasErrors ? '⚠️' : '🎉'}</div>
          <div className={styles.resultInfo}>
            <div className={styles.resultTitle}>
              {hasErrors
                ? `Import completed with ${importResult.failed} error${importResult.failed !== 1 ? 's' : ''}`
                : 'Import successful!'}
            </div>
            <div className={styles.resultDetail}>
              {importResult.succeeded} task{importResult.succeeded !== 1 ? 's' : ''} added to <strong>{project.title}</strong>
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
            text="View Imported Tasks"
            onClick={() => { onImportComplete(); onDismiss(); }}
          />
          <DefaultButton text="Close" onClick={onDismiss} />
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
          <DefaultButton text="Back" onClick={() => setStep('source')} />
          <DefaultButton text="Cancel" onClick={onDismiss} />
        </Stack>
      );
    }

    // source step
    const needsMap = importSource?.needsMapping || sourceType === 'excel';
    return (
      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <PrimaryButton
          text={needsMap ? 'Next: Map Columns →' : `Import ${taskCount} Tasks`}
          disabled={!canProceedFromSource() || taskCount === 0}
          onClick={handleNext}
        />
        <DefaultButton text="Cancel" onClick={onDismiss} />
      </Stack>
    );
  };

  // ─── Step indicator ────────────────────────────────────────────────────────

  const steps: Array<{ id: ImportStep; label: string }> = [
    { id: 'source', label: 'Source' },
    { id: 'map', label: 'Map' },
    { id: 'importing', label: 'Import' },
    { id: 'done', label: 'Done' },
  ];

  const stepOrder: ImportStep[] = ['source', 'map', 'importing', 'done'];
  const currentIdx = stepOrder.indexOf(step);

  return (
    <Panel
      isOpen={isOpen}
      type={PanelType.medium}
      headerText={`Import Tasks into "${project.title}"`}
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
        {step === 'source' && renderSourceStep()}
        {step === 'map' && renderMapStep()}
        {step === 'importing' && renderImportingStep()}
        {step === 'done' && renderDoneStep()}
      </div>
    </Panel>
  );
};

export default ImportPanel;
