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
}

const commit = (e: React.KeyboardEvent, onChange: () => void): void => {
  if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onChange(); }
};

// Shared by TaskPanel (custom bar color, with an "Auto" option) and
// ProjectPanel (project color, always set) — was previously duplicated
// almost verbatim in both. Swatches are keyboard/screen-reader accessible
// (role="radio" + arrow-free tab navigation, matching a color swatch's
// usual single-select semantics without requiring roving tabindex).
export const ColorSwatchPicker: React.FC<IColorSwatchPickerProps> = ({
  colors, value, onChange, size = 28, allowAuto,
}) => {
  const isCustom = !!value && !colors.includes(value);

  return (
    <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 6, alignItems: 'center' }} role="radiogroup">
      {allowAuto && (
        <div
          role="radio"
          aria-checked={!value}
          aria-label={strings.ColorSwatch_AutoColorAriaLabel}
          tabIndex={0}
          onClick={() => onChange('')}
          onKeyDown={e => commit(e, () => onChange(''))}
          title={strings.ColorSwatch_AutoColorAriaLabel}
          style={{
            width: size, height: size, borderRadius: '50%',
            background: 'linear-gradient(135deg, #ccc 50%, #fff 50%)',
            cursor: 'pointer',
            border: !value ? '3px solid #323130' : '3px solid transparent',
            boxSizing: 'border-box',
          }}
        />
      )}
      {colors.map(c => (
        <div
          key={c}
          role="radio"
          aria-checked={value === c}
          aria-label={c}
          tabIndex={0}
          onClick={() => onChange(c)}
          onKeyDown={e => commit(e, () => onChange(c))}
          style={{
            width: size, height: size, borderRadius: '50%', background: c,
            cursor: 'pointer',
            border: value === c ? '3px solid #323130' : '3px solid transparent',
            outline: value === c ? `2px solid ${c}` : 'none',
            outlineOffset: 2,
            boxSizing: 'border-box',
          }}
        />
      ))}
      <label
        title={strings.ColorSwatch_PickCustomColorTitle}
        style={{ position: 'relative', width: size, height: size, cursor: 'pointer', flexShrink: 0 }}
      >
        <div
          aria-hidden="true"
          style={{
            width: size, height: size, borderRadius: '50%',
            background: isCustom ? value : 'conic-gradient(red,yellow,lime,cyan,blue,magenta,red)',
            border: isCustom ? '3px solid #323130' : '2px solid #EDEBE9',
            outline: isCustom ? `2px solid ${value}` : 'none',
            outlineOffset: 2, boxSizing: 'border-box',
          }}
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
