import * as React from 'react';
import { Panel, PanelType, Toggle } from '@fluentui/react';
import {
  IGanttDisplaySettings, GanttColorBy, GanttWeekLabel,
  GanttHeaderTheme, GanttBarStyle, HEADER_THEME_COLORS,
} from '../../models';
import * as strings from 'SmartGanttWebPartStrings';
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
    { id: 'status',   label: strings.GanttSettings_ByStatus,   icon: '⬛' },
    { id: 'priority', label: strings.GanttSettings_ByPriority, icon: '🔺' },
    { id: 'phase',    label: strings.GanttSettings_ByPhase,    icon: '🏷' },
    { id: 'health',   label: strings.GanttSettings_ByHealth,   icon: '❤' },
  ];

  const weekOptions: { id: GanttWeekLabel; label: string; desc: string }[] = [
    { id: 'dates',   label: strings.GanttSettings_WeekLabelDates,   desc: strings.GanttSettings_WeekLabelDatesDesc },
    { id: 'project', label: strings.GanttSettings_WeekLabelProject, desc: strings.GanttSettings_WeekLabelProjectDesc },
    { id: 'iso',     label: strings.GanttSettings_WeekLabelIso,     desc: strings.GanttSettings_WeekLabelIsoDesc },
  ];

  const barOptions: { id: GanttBarStyle; label: string }[] = [
    { id: 'gradient', label: strings.GanttSettings_BarStyleGradient },
    { id: 'flat',     label: strings.GanttSettings_BarStyleFlat },
  ];

  const themes: { id: GanttHeaderTheme; label: string }[] = [
    { id: 'dark',   label: strings.GanttSettings_ThemeDark },
    { id: 'navy',   label: strings.GanttSettings_ThemeNavy },
    { id: 'teal',   label: strings.GanttSettings_ThemeTeal },
    { id: 'purple', label: strings.GanttSettings_ThemePurple },
    { id: 'light',  label: strings.GanttSettings_ThemeLight },
  ];

  const heights: { value: number; label: string; barH: number }[] = [
    { value: 36, label: strings.GanttSettings_HeightCompact,  barH: 3 },
    { value: 40, label: strings.GanttSettings_HeightNormal,   barH: 5 },
    { value: 52, label: strings.GanttSettings_HeightSpacious, barH: 7 },
  ];

  return (
    <Panel
      isOpen={isOpen}
      type={PanelType.smallFixedFar}
      headerText={strings.GanttSettings_PanelHeader}
      onDismiss={onDismiss}
      isLightDismiss
      isBlocking={false}
    >
      <div className={styles.panel}>

        {/* ── Color coding ─────────────────────────────────────────────── */}
        <div className={styles.section}>
          <div className={styles.sectionTitle}>{strings.GanttSettings_ColorCodingSection}</div>
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
          <div className={styles.sectionTitle}>{strings.GanttSettings_HeaderColorSection}</div>
          <div className={styles.themeGrid}>
            {themes.map(t => (
              <button
                key={t.id}
                type="button"
                className={`${styles.themeSwatch} ${settings.headerTheme === t.id ? styles.selected : ''} ${t.id === 'light' ? styles.light : ''}`}
                style={{ background: HEADER_THEME_COLORS[t.id].bg }}
                title={t.label}
                aria-label={t.label}
                aria-pressed={settings.headerTheme === t.id}
                onClick={() => set('headerTheme', t.id)}
              >
                {t.label[0]}
              </button>
            ))}
          </div>
        </div>

        {/* ── Header labels ─────────────────────────────────────────────── */}
        <div className={styles.section}>
          <div className={styles.sectionTitle}>{strings.GanttSettings_HeaderLabelsSection}</div>
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
              {strings.GanttSettings_ProjectWeeksHint}
            </div>
          )}
        </div>

        {/* ── Bar style ─────────────────────────────────────────────────── */}
        <div className={styles.section}>
          <div className={styles.sectionTitle}>{strings.GanttSettings_BarStyleSection}</div>
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
          <div className={styles.sectionTitle}>{strings.GanttSettings_RowHeightSection}</div>
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
          <div className={styles.sectionTitle}>{strings.GanttSettings_ShowHideSection}</div>
          {([['showWeekends', strings.GanttSettings_WeekendShading]] as const).map(([key, label]) => (
            <div key={key} className={styles.toggleRow}>
              <span className={styles.toggleLabel}>{label}</span>
              <Toggle checked={settings[key]} onChange={(_, v) => set(key, !!v)} ariaLabel={label} styles={{ root: { margin: 0 } }} />
            </div>
          ))}
          <div className={styles.toggleRow}>
            <span className={styles.toggleLabel}>{strings.GanttSettings_DependencyArrows}</span>
            <Toggle checked={settings.showDependencies} onChange={(_, v) => set('showDependencies', !!v)} ariaLabel={strings.GanttSettings_DependencyArrows} styles={{ root: { margin: 0 } }} />
          </div>
          {settings.showDependencies && (
            <>
              {([
                ['showCriticalPathOnly', strings.GanttSettings_CriticalPathAlwaysVisible],
                ['dependenciesOnHover',  strings.GanttSettings_AllOthersOnHoverOnly],
              ] as const).map(([key, label]) => (
                <div key={key} className={styles.toggleRow} style={{ paddingLeft: 20 }}>
                  <span className={styles.toggleLabel} style={{ color: '#605E5C' }}>{label}</span>
                  <Toggle checked={settings[key]} onChange={(_, v) => set(key, !!v)} ariaLabel={label} styles={{ root: { margin: 0 } }} />
                </div>
              ))}
            </>
          )}
          {([
            ['showCriticalPath', strings.GanttSettings_CriticalPathHighlight],
            ['showProgressText', strings.GanttSettings_ProgressPercentOnBars],
            ['showAssignee',     strings.GanttSettings_AssigneeNameOnBars],
            ['showHealthBadges', strings.GanttSettings_HealthStatusBadges],
          ] as const).map(([key, label]) => (
            <div key={key} className={styles.toggleRow}>
              <span className={styles.toggleLabel}>{label}</span>
              <Toggle checked={settings[key] as boolean} onChange={(_, v) => set(key, !!v)} ariaLabel={label} styles={{ root: { margin: 0 } }} />
            </div>
          ))}
        </div>

      </div>
    </Panel>
  );
};

export default GanttSettings;
