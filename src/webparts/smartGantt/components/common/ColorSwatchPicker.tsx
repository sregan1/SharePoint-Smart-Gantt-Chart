import * as React from 'react';
import * as strings from 'SmartGanttWebPartStrings';

export interface IColorSwatchPickerProps {
  colors: string[];
  /** Empty string means "no color selected" — only meaningful when allowAuto is set. */
  value: string;
  onChange: (color: string) => void;
  size?: number;
  /** Renders a leading "Auto" swatch that sets value back to ''. */
  allowAuto?: boolean;
  /** Accessible name for the whole group (e.g. the field label). */
  ariaLabel?: string;
}

const commit = (e: React.KeyboardEvent, onChange: () => void): void => {
  if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onChange(); }
};

const CHECK_STYLE = {
  color: '#fff',
  fontSize: 13,
  fontWeight: 700,
  lineHeight: 1,
  textShadow: '0 0 2px rgba(0,0,0,0.7)',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  width: '100%',
  height: '100%',
  // Windows high-contrast mode would otherwise repaint the swatch fill.
  forcedColorAdjust: 'none',
} as React.CSSProperties;

// Shared by TaskPanel (custom bar color, with an "Auto" option) and
// ProjectPanel (project color, always set) — was previously duplicated
// almost verbatim in both. Swatches follow the radiogroup pattern: one tab
// stop (the selected swatch), arrow keys move the selection, and the selected
// swatch carries a checkmark so it isn't indicated by border color alone.
export const ColorSwatchPicker: React.FC<IColorSwatchPickerProps> = ({
  colors, value, onChange, size = 28, allowAuto, ariaLabel,
}) => {
  const isCustom = !!value && !colors.includes(value);
  const swatchRefs = React.useRef<Array<HTMLDivElement | null>>([]);

  // Selectable values in visual order (the custom-color picker is its own control).
  const values: string[] = allowAuto ? ['', ...colors] : [...colors];
  const selectedIdx = values.indexOf(value);
  // Custom color selected (or nothing matches): the first swatch stays reachable.
  const tabStopIdx = selectedIdx >= 0 ? selectedIdx : 0;

  const handleArrow = (e: React.KeyboardEvent, idx: number): void => {
    let next = -1;
    if (e.key === 'ArrowRight' || e.key === 'ArrowDown') next = (idx + 1) % values.length;
    else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') next = (idx - 1 + values.length) % values.length;
    else if (e.key === 'Home') next = 0;
    else if (e.key === 'End') next = values.length - 1;
    if (next < 0) return;
    e.preventDefault();
    onChange(values[next]);
    swatchRefs.current[next]?.focus();
  };

  const swatchIdx = (v: string): number => values.indexOf(v);

  return (
    <div
      style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 6, alignItems: 'center' }}
      role="radiogroup"
      aria-label={ariaLabel}
    >
      {allowAuto && (
        <div
          ref={el => { swatchRefs.current[0] = el; }}
          role="radio"
          aria-checked={!value}
          aria-label={strings.ColorSwatch_AutoColorAriaLabel}
          tabIndex={tabStopIdx === 0 ? 0 : -1}
          onClick={() => onChange('')}
          onKeyDown={e => { commit(e, () => onChange('')); handleArrow(e, 0); }}
          title={strings.ColorSwatch_AutoColorAriaLabel}
          style={{
            width: size, height: size, borderRadius: '50%',
            background: 'linear-gradient(135deg, #ccc 50%, #fff 50%)',
            cursor: 'pointer',
            border: !value ? '3px solid var(--neutralPrimary, #323130)' : '3px solid transparent',
            boxSizing: 'border-box',
          }}
        />
      )}
      {colors.map(c => {
        const idx = swatchIdx(c);
        const selected = value === c;
        return (
          <div
            key={c}
            ref={el => { swatchRefs.current[idx] = el; }}
            role="radio"
            aria-checked={selected}
            aria-label={c}
            tabIndex={tabStopIdx === idx ? 0 : -1}
            onClick={() => onChange(c)}
            onKeyDown={e => { commit(e, () => onChange(c)); handleArrow(e, idx); }}
            style={{
              width: size, height: size, borderRadius: '50%', background: c,
              cursor: 'pointer',
              border: selected ? '3px solid var(--neutralPrimary, #323130)' : '3px solid transparent',
              outline: selected ? `2px solid ${c}` : 'none',
              outlineOffset: 2,
              boxSizing: 'border-box',
              forcedColorAdjust: 'none',
            } as React.CSSProperties}
          >
            {selected && <span aria-hidden="true" style={CHECK_STYLE}>✓</span>}
          </div>
        );
      })}
      <label
        title={strings.ColorSwatch_PickCustomColorTitle}
        style={{ position: 'relative', width: size, height: size, cursor: 'pointer', flexShrink: 0 }}
      >
        <div
          aria-hidden="true"
          style={{
            width: size, height: size, borderRadius: '50%',
            background: isCustom ? value : 'conic-gradient(red,yellow,lime,cyan,blue,magenta,red)',
            border: isCustom ? '3px solid var(--neutralPrimary, #323130)' : '2px solid var(--neutralLight, #EDEBE9)',
            outline: isCustom ? `2px solid ${value}` : 'none',
            outlineOffset: 2, boxSizing: 'border-box',
            forcedColorAdjust: 'none',
          } as React.CSSProperties}
        />
        <input
          type="color"
          value={isCustom ? value : '#0078D4'}
          onChange={e => onChange(e.target.value)}
          aria-label={strings.ColorSwatch_CustomColorAriaLabel}
          style={{ position: 'absolute', opacity: 0, width: 0, height: 0, pointerEvents: 'none' }}
        />
      </label>
    </div>
  );
};

export default ColorSwatchPicker;
