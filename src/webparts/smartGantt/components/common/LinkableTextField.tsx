import * as React from 'react';
import { TextField, ITextField, Label, ActionButton } from '@fluentui/react';
import * as strings from 'SmartGanttWebPartStrings';
import { THEME as T } from './theme';
import { Linkify, containsUrl } from './Linkify';

interface ILinkableTextFieldProps {
  label: string;
  value: string;
  onChange: (value: string) => void;
  rows: number;
  placeholder?: string;
}

const LINE_HEIGHT = 20;

/**
 * Multiline text field whose URLs are clickable. A text box can't hold links,
 * so text that contains a URL is shown as read-only linkified text until the
 * user clicks it (or presses Edit); leaving the box returns to the read-only
 * view. Text without a URL is always a plain editable box, so typing or
 * pasting a link never switches modes mid-edit.
 */
export const LinkableTextField: React.FC<ILinkableTextFieldProps> = ({
  label, value, onChange, rows, placeholder,
}) => {
  const hasUrl = containsUrl(value);
  const [editing, setEditing] = React.useState(!hasUrl);
  const fieldRef = React.useRef<ITextField>(null);
  const focusOnEdit = React.useRef(false);

  React.useEffect(() => {
    if (editing && focusOnEdit.current) {
      focusOnEdit.current = false;
      if (fieldRef.current) fieldRef.current.focus();
    }
  }, [editing]);

  const startEdit = (): void => {
    focusOnEdit.current = true;
    setEditing(true);
  };

  if (editing || !hasUrl) {
    return (
      <TextField
        componentRef={fieldRef}
        label={label}
        value={value}
        onChange={(_, v) => onChange(v || '')}
        onBlur={() => { if (hasUrl) setEditing(false); }}
        multiline
        rows={rows}
        resizable={false}
        placeholder={placeholder}
      />
    );
  }

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <Label>{label}</Label>
        <ActionButton iconProps={{ iconName: 'Edit' }} onClick={startEdit} styles={{ root: { height: 24 } }}>
          {strings.ListView_EditTitle}
        </ActionButton>
      </div>
      <div
        onClick={(e) => { if (!(e.target as HTMLElement).closest('a')) startEdit(); }}
        style={{
          minHeight: rows * LINE_HEIGHT, padding: '5px 8px', fontSize: 14, lineHeight: `${LINE_HEIGHT}px`,
          border: `1px solid ${T.border}`, borderRadius: 2, background: T.surface,
          cursor: 'text', overflowWrap: 'anywhere',
        }}
      >
        <Linkify text={value} />
      </div>
    </div>
  );
};
