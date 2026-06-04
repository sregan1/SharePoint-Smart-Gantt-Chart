import * as React from 'react';
import { IImportSource, IMPORTABLE_FIELDS, ColumnMapping, ImportableField } from '../../services/ImportService';
import styles from './ColumnMapper.module.scss';

interface IColumnMapperProps {
  source: IImportSource;
  mapping: ColumnMapping;
  onChange: (mapping: ColumnMapping) => void;
}

const PREVIEW_ROWS = 3;
const PREVIEW_COLS = 5;

export const ColumnMapper: React.FC<IColumnMapperProps> = ({ source, mapping, onChange }) => {
  const hasTitleMapped = Object.values(mapping).includes('title');
  const skippedCount = Object.values(mapping).filter(v => v === 'skip').length;

  // Columns to preview: only those that are mapped (not skipped), up to PREVIEW_COLS
  const previewCols = source.headers
    .filter(h => mapping[h] && mapping[h] !== 'skip')
    .slice(0, PREVIEW_COLS);

  const previewRows = source.rows.slice(0, PREVIEW_ROWS);

  const handleChange = (header: string, value: ImportableField): void => {
    // Ensure the same target field isn't mapped twice (except 'skip')
    const newMapping = { ...mapping };
    if (value !== 'skip') {
      Object.keys(newMapping).forEach(h => {
        if (newMapping[h] === value && h !== header) {
          newMapping[h] = 'skip';
        }
      });
    }
    newMapping[header] = value;
    onChange(newMapping);
  };

  const fieldLabel = (field: ImportableField): string => {
    const def = IMPORTABLE_FIELDS.find(f => f.key === field);
    return def ? def.label : field;
  };

  return (
    <div className={styles.mapper}>
      {!hasTitleMapped && (
        <div className={styles.warningBanner}>
          <span className={styles.warningIcon}>⚠</span>
          <span><strong>Task Name</strong> is required — map it to a source column before importing.</span>
        </div>
      )}

      {hasTitleMapped && skippedCount > 0 && (
        <div className={styles.infoBanner}>
          <span className={styles.infoIcon}>ℹ</span>
          <span><strong>{skippedCount}</strong> column{skippedCount !== 1 ? 's' : ''} will be skipped. Review the mappings below and adjust if needed.</span>
        </div>
      )}

      {/* Header row */}
      <div className={styles.mapperHeader}>
        <span className={styles.mapperHeaderLabel}>Source Column</span>
        <span />
        <span className={styles.mapperHeaderLabel}>Smart Gantt Field</span>
        <span className={styles.mapperHeaderLabel}>Status</span>
      </div>

      {/* Mapping rows */}
      {source.headers.map(header => {
        const mapped = mapping[header];
        const isSkipped = mapped === 'skip';
        const wasAutoMapped = source.autoMapping[header] !== 'skip' && source.autoMapping[header] === mapped;

        return (
          <div
            key={header}
            className={`${styles.mapperRow} ${isSkipped ? styles.skipped : ''} ${wasAutoMapped ? styles.autoMapped : ''}`}
          >
            <span className={styles.sourceCol} title={header}>{header}</span>
            <span className={styles.arrow}>→</span>
            <select
              className={styles.targetSelect}
              value={mapped || 'skip'}
              onChange={e => handleChange(header, e.target.value as ImportableField)}
            >
              {IMPORTABLE_FIELDS.map(f => (
                <option key={f.key} value={f.key}>
                  {f.label}{f.required ? ' *' : ''}
                </option>
              ))}
            </select>
            <span>
              {isSkipped ? (
                <span className={styles.skippedBadge}>Skipped</span>
              ) : (
                <span className={styles.mappedBadge}>
                  {wasAutoMapped ? '✓ Auto' : '✓ Set'}
                </span>
              )}
            </span>
          </div>
        );
      })}

      {/* Preview */}
      {previewCols.length > 0 && previewRows.length > 0 && (
        <div className={styles.previewSection}>
          <div className={styles.previewTitle}>Preview (first {Math.min(PREVIEW_ROWS, previewRows.length)} rows)</div>
          <div style={{ overflowX: 'auto' }}>
            <table className={styles.previewTable}>
              <thead>
                <tr>
                  {previewCols.map(col => (
                    <th key={col} title={col}>{fieldLabel(mapping[col])}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {previewRows.map((row, i) => (
                  <tr key={i}>
                    {previewCols.map(col => (
                      <td key={col} title={row[col]}>{row[col] || '—'}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
};

export default ColumnMapper;
