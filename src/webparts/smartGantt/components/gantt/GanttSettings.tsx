import * as React from 'react';
import { Panel, PanelType, Toggle } from '@fluentui/react';
import {
  IGanttDisplaySettings, GanttColorBy, GanttWeekLabel,
  GanttHeaderTheme, GanttBarStyle, HEADER_THEME_COLORS,
} from '../../models';
import styles from './GanttSettings.module.scss';

interface IGanttSettingsProps {
  isOpen: boolean;
  settings: IGanttDisplaySettings;
  onChange: (settings: IGanttDisplaySettings) => void;
  onDismiss: () => void;
}

export const GanttSettings: React.FC<IGanttSettingsProps> = ({
  isOpen, settings, onChange, onDismiss,
}) => {
  const set = <K extends keyof IGanttDisplaySettings>(key: K, value: IGanttDisplaySettings[K]): void => {
    onChange({ ...settings, [key]: value });
  };

  const colorOptions: { id: GanttColorBy; label: string; icon: string }[] = [
    { id: 'status',   label: 'By Status',   icon: '⬛' },
    { id: 'priority', label: 'By Priority', icon: '🔺' },
    { id: 'phase',    label: 'By Phase',    icon: '🏷' },
    { id: 'health',   label: 'By Health',   icon: '❤' },
  ];

  const weekOptions: { id: GanttWeekLabel; label: string; desc: string }[] = [
    { id: 'dates',   label: 'Dates',          desc: 'Jun 3, Jun 10…' },
    { id: 'project', label: 'Project Weeks',  desc: 'W1, W2…' },
    { id: 'iso',     label: 'Calendar Weeks', desc: 'W23, W24…' },
  ];

  const barOptions: { id: GanttBarStyle; label: string }[] = [
    { id: 'gradient', label: 'Gradient' },
    { id: 'flat',     label: 'Flat' },
  ];

  const themes: { id: GanttHeaderTheme; label: string }[] = [
    { id: 'dark',   label: 'Dark' },
    { id: 'navy',   label: 'Navy' },
    { id: 'teal',   label: 'Teal' },
    { id: 'purple', label: 'Purple' },
    { id: 'light',  label: 'Light' },
  ];

  const heights: { value: number; label: string; barH: number }[] = [
    { value: 32, label: 'Compact',  barH: 3 },
    { value: 40, label: 'Normal',   barH: 5 },
    { value: 52, label: 'Spacious', barH: 7 },
  ];

  return (
    <Panel
      isOpen={isOpen}
      type={PanelType.smallFixedFar}
      headerText="⚙ Options"
      onDismiss={onDismiss}
      isLightDismiss
      isBlocking={false}
    >
      <div className={styles.panel}>

        {/* ── Color coding ─────────────────────────────────────────────── */}
        <div className={styles.section}>
          <div className={styles.sectionTitle}>Color Coding</div>
          <div className={styles.chipGroup}>
            {colorOptions.map(o => (
              <button
                key={o.id}
                className={`${styles.chip} ${settings.colorBy === o.id ? styles.selected : ''}`}
                onClick={() => set('colorBy', o.id)}
              >
                <span>{o.icon}</span> {o.label}
              </button>
            ))}
          </div>
        </div>

        {/* ── Header theme ──────────────────────────────────────────────── */}
        <div className={styles.section}>
          <div className={styles.sectionTitle}>Header Color</div>
          <div className={styles.themeGrid}>
            {themes.map(t => (
              <div
                key={t.id}
                className={`${styles.themeSwatch} ${settings.headerTheme === t.id ? styles.selected : ''} ${t.id === 'light' ? styles.light : ''}`}
                style={{ background: HEADER_THEME_COLORS[t.id].bg }}
                title={t.label}
                onClick={() => set('headerTheme', t.id)}
              >
                {t.label[0]}
              </div>
            ))}
          </div>
        </div>

        {/* ── Header labels ─────────────────────────────────────────────── */}
        <div className={styles.section}>
          <div className={styles.sectionTitle}>Header Labels</div>
          <div className={styles.chipGroup}>
            {weekOptions.map(o => (
              <button
                key={o.id}
                className={`${styles.chip} ${settings.weekLabel === o.id ? styles.selected : ''}`}
                onClick={() => set('weekLabel', o.id)}
                title={o.desc}
              >
                {o.label}
                <span style={{ fontWeight: 400, opacity: 0.75 }}>&ensp;({o.desc})</span>
              </button>
            ))}
          </div>
          {settings.weekLabel === 'project' && (
            <div style={{ fontSize: 12, color: '#605E5C', marginTop: 4 }}>
              Week 1 starts from the Monday of your earliest task — great for presentations.
            </div>
          )}
        </div>

        {/* ── Bar style ─────────────────────────────────────────────────── */}
        <div className={styles.section}>
          <div className={styles.sectionTitle}>Bar Style</div>
          <div className={styles.chipGroup}>
            {barOptions.map(o => (
              <button
                key={o.id}
                className={`${styles.chip} ${settings.barStyle === o.id ? styles.selected : ''}`}
                onClick={() => set('barStyle', o.id)}
              >
                {o.label}
              </button>
            ))}
          </div>
        </div>

        {/* ── Row height ────────────────────────────────────────────────── */}
        <div className={styles.section}>
          <div className={styles.sectionTitle}>Row Height</div>
          <div className={styles.heightGroup}>
            {heights.map(h => (
              <button
                key={h.value}
                className={`${styles.heightBtn} ${settings.rowHeight === h.value ? styles.selected : ''}`}
                onClick={() => set('rowHeight', h.value)}
              >
                <div className={styles.heightPreview}>
                  {[...Array(3)].map((_, i) => (
                    <div key={i} className={styles.heightBar} style={{ height: h.barH }} />
                  ))}
                </div>
                {h.label}
              </button>
            ))}
          </div>
        </div>

        {/* ── Toggles ───────────────────────────────────────────────────── */}
        <div className={styles.section}>
          <div className={styles.sectionTitle}>Show / Hide</div>
          {([
            ['showWeekends',     'Weekend shading'],
            ['showDependencies', 'Dependency arrows'],
            ['showProgressText', 'Progress % on bars'],
            ['showAssignee',     'Assignee name on bars'],
            ['showHealthBadges', 'Health status badges'],
          ] as const).map(([key, label]) => (
            <div key={key} className={styles.toggleRow}>
              <span className={styles.toggleLabel}>{label}</span>
              <Toggle
                checked={settings[key] as boolean}
                onChange={(_, v) => set(key, !!v)}
                styles={{ root: { margin: 0 } }}
              />
            </div>
          ))}
        </div>

      </div>
    </Panel>
  );
};

export default GanttSettings;
