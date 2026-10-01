import * as React from 'react';
import { Label, TextField } from '@fluentui/react';
import * as strings from 'SmartGanttWebPartStrings';
import { formatString } from '../localeUtils';
import { ITask } from '../../models';
import { initials, stringToColor } from '../../utils/taskDisplayUtils';
import { VISUALLY_HIDDEN, useUniqueId } from './a11y';
import styles from './PeoplePicker.module.scss';

export interface IPersonSuggestion {
  name: string;
  email: string;
}

/** Anything that can search the directory — SharePointService satisfies this. */
export interface IPeopleSearchProvider {
  searchPeople: (query: string) => Promise<IPersonSuggestion[]>;
}

export interface IPeoplePickerProps {
  label?: string;
  required?: boolean;
  /** Stored free-text display name (ITask.assignedTo / IProject.projectManager). */
  name: string;
  /** Optional stored email (ITask.assignedToEmail / IProject.projectManagerEmail). */
  email: string;
  onChange: (name: string, email: string) => void;
  /** Previously used people, offered as suggestions alongside directory matches. */
  recent?: IPersonSuggestion[];
  /** When omitted the picker falls back to suggestions-only behavior. */
  spService?: IPeopleSearchProvider;
  placeholder?: string;
  /** Validation message (e.g. an invalid email) rendered under the control. */
  errorMessage?: string;
}

type OptionKind = 'directory' | 'recent' | 'external';
interface IOption {
  id: string;
  kind: OptionKind;
  name: string;
  email: string;
}

const SEARCH_DEBOUNCE_MS = 250;
const MAX_RECENT = 6;
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

/** Empty is valid (email is optional). */
export function isValidEmail(value: string): boolean {
  return !value.trim() || EMAIL_RE.test(value.trim());
}

/**
 * Distinct people previously used as assignees, with the email last stored for
 * each name. `names` (e.g. the panel's knownUsers list) keeps its order first.
 */
export function buildRecentPeople(names: string[], tasks: ITask[]): IPersonSuggestion[] {
  const emailByName = new Map<string, string>();
  tasks.forEach(t => {
    if (t.assignedTo && t.assignedToEmail && !emailByName.has(t.assignedTo)) {
      emailByName.set(t.assignedTo, t.assignedToEmail);
    }
  });
  const seen = new Set<string>();
  const result: IPersonSuggestion[] = [];
  const add = (n: string): void => {
    const key = n.trim().toLowerCase();
    if (!key || seen.has(key)) return;
    seen.add(key);
    result.push({ name: n, email: emailByName.get(n) || '' });
  };
  names.forEach(add);
  tasks.forEach(t => add(t.assignedTo || ''));
  return result;
}

// Assignee stays a stored free-text string (+ optional email): the combobox
// only helps fill it in. Directory matches set both name and email; the
// always-present "external" option accepts whatever was typed, so consultants
// who aren't in the directory can still be assigned.
export const PeoplePicker: React.FC<IPeoplePickerProps> = ({
  label, required, name, email, onChange, recent = [], spService, placeholder, errorMessage,
}) => {
  const inputId = useUniqueId('pp-input');
  const listId = useUniqueId('pp-list');
  const wrapRef = React.useRef<HTMLDivElement>(null);
  const inputRef = React.useRef<HTMLInputElement>(null);

  const [open, setOpen] = React.useState(false);
  const [active, setActive] = React.useState(-1);
  const [results, setResults] = React.useState<IPersonSuggestion[]>([]);
  const [searching, setSearching] = React.useState(false);
  // null = derive from the stored value (a name without an email is treated as
  // an external entry); true/false once the user has typed or chosen.
  const [externalFlag, setExternalFlag] = React.useState<boolean | null>(null);

  const timerRef = React.useRef<number | undefined>(undefined);
  const seqRef = React.useRef(0);
  const mountedRef = React.useRef(true);

  React.useEffect(() => {
    mountedRef.current = true;
    return () => {
      mountedRef.current = false;
      seqRef.current++;
      if (timerRef.current !== undefined) window.clearTimeout(timerRef.current);
    };
  }, []);

  // A cleared/reset value (new task opened) forgets the earlier external choice.
  React.useEffect(() => {
    if (!name) setExternalFlag(null);
  }, [name]);

  // Close on outside click
  React.useEffect(() => {
    const handler = (e: MouseEvent): void => {
      if (wrapRef.current && !wrapRef.current.contains(e.target as Node)) {
        setOpen(false);
        setActive(-1);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const cancelSearch = (): void => {
    seqRef.current++;
    if (timerRef.current !== undefined) window.clearTimeout(timerRef.current);
    timerRef.current = undefined;
  };

  // Debounced directory search. Each call bumps a sequence number so a slow
  // response for an older query can never overwrite a newer one.
  const scheduleSearch = (query: string): void => {
    cancelSearch();
    const q = query.trim();
    if (!spService || !q) {
      setResults([]);
      setSearching(false);
      return;
    }
    const seq = seqRef.current;
    setSearching(true);
    timerRef.current = window.setTimeout(() => {
      spService.searchPeople(q)
        .then(found => {
          if (!mountedRef.current || seq !== seqRef.current) return;
          setResults(found || []);
          setSearching(false);
        })
        .catch(() => {
          if (!mountedRef.current || seq !== seqRef.current) return;
          setResults([]);
          setSearching(false);
        });
    }, SEARCH_DEBOUNCE_MS);
  };

  const query = name.trim();
  const queryLower = query.toLowerCase();

  const { options, directoryCount, recentCount } = React.useMemo(() => {
    const dir: IOption[] = results.map((p, i) => ({ id: `d${i}`, kind: 'directory', name: p.name, email: p.email }));
    const dirKeys = new Set(results.map(p => `${p.name}|${p.email}`.toLowerCase()));
    const rec: IOption[] = recent
      .filter(p => p.name !== name || p.email !== email)
      .filter(p => !queryLower
        || p.name.toLowerCase().indexOf(queryLower) !== -1
        || p.email.toLowerCase().indexOf(queryLower) !== -1)
      .filter(p => !dirKeys.has(`${p.name}|${p.email}`.toLowerCase()))
      .slice(0, MAX_RECENT)
      .map((p, i) => ({ id: `r${i}`, kind: 'recent', name: p.name, email: p.email }));
    const ext: IOption[] = query
      ? [{ id: 'x', kind: 'external', name: query, email: '' }]
      : [];
    return { options: [...dir, ...rec, ...ext], directoryCount: dir.length, recentCount: rec.length };
  }, [results, recent, query, queryLower, name, email]);

  const choose = (opt: IOption): void => {
    cancelSearch();
    setSearching(false);
    if (opt.kind === 'external') {
      onChange(opt.name, '');
      setExternalFlag(true);
    } else {
      onChange(opt.name, opt.email);
      // A recent entry with no stored email is still a free-text name.
      setExternalFlag(!opt.email);
    }
    setOpen(false);
    setActive(-1);
    inputRef.current?.focus();
  };

  const handleInput = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const text = e.target.value;
    // Typing invalidates any email that came from a previously chosen person.
    onChange(text, '');
    setExternalFlag(true);
    setOpen(true);
    setActive(-1);
    scheduleSearch(text);
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      if (!open) { setOpen(true); setActive(options.length > 0 ? 0 : -1); return; }
      setActive(a => Math.min(a + 1, options.length - 1));
    } else if (e.key === 'ArrowUp') {
      if (!open) return;
      e.preventDefault();
      setActive(a => Math.max(a - 1, 0));
    } else if (e.key === 'Enter') {
      if (open && active >= 0 && options[active]) {
        e.preventDefault();
        choose(options[active]);
      }
    } else if (e.key === 'Escape') {
      if (open) {
        // Keep Escape from also dismissing the surrounding Panel.
        e.stopPropagation();
        e.nativeEvent.stopImmediatePropagation();
        setOpen(false);
        setActive(-1);
      }
    } else if (e.key === 'Tab') {
      setOpen(false);
      setActive(-1);
    }
  };

  const isExternal = externalFlag === null ? !email : externalFlag;
  const showEmailField = !!query && isExternal;
  const listVisible = open && options.length > 0;
  const activeId = listVisible && active >= 0 && options[active] ? `${listId}-${options[active].id}` : undefined;
  const emailInvalid = !isValidEmail(email);
  const errorId = `${inputId}-err`;

  const renderOption = (opt: IOption): React.ReactNode => (
    <div
      key={opt.id}
      id={`${listId}-${opt.id}`}
      role="option"
      aria-selected={options[active]?.id === opt.id}
      className={`${styles.option} ${options[active]?.id === opt.id ? styles.active : ''}`}
      onMouseDown={e => { e.preventDefault(); choose(opt); }}
      onMouseMove={() => { const idx = options.indexOf(opt); if (idx !== active) setActive(idx); }}
    >
      {opt.kind === 'external' ? (
        <span className={styles.externalIcon} aria-hidden="true">+</span>
      ) : (
        <span className={styles.avatar} style={{ background: stringToColor(opt.name) }} aria-hidden="true">
          {initials(opt.name)}
        </span>
      )}
      <span className={styles.optionText}>
        <span className={styles.optionName}>
          {opt.kind === 'external'
            ? formatString(strings.Common_PeoplePicker_UseExternal, { name: opt.name })
            : opt.name}
        </span>
        {opt.kind !== 'external' && opt.email && <span className={styles.optionEmail}>{opt.email}</span>}
      </span>
    </div>
  );

  const directoryOptions = options.slice(0, directoryCount);
  const recentOptions = options.slice(directoryCount, directoryCount + recentCount);
  const externalOptions = options.slice(directoryCount + recentCount);

  return (
    <div className={styles.wrap} ref={wrapRef}>
      {label && <Label htmlFor={inputId} required={required}>{label}</Label>}
      <div className={styles.inputRow}>
        <input
          ref={inputRef}
          id={inputId}
          className={`${styles.input} ${errorMessage || emailInvalid ? styles.inputError : ''}`}
          value={name}
          placeholder={placeholder}
          onChange={handleInput}
          onFocus={() => setOpen(true)}
          onKeyDown={handleKeyDown}
          autoComplete="off"
          role="combobox"
          aria-autocomplete="list"
          aria-haspopup="listbox"
          aria-expanded={listVisible}
          aria-controls={listId}
          aria-activedescendant={activeId}
          aria-required={required}
          aria-invalid={!!errorMessage}
          aria-describedby={errorMessage ? errorId : undefined}
        />
        {name && (
          <button
            type="button"
            className={styles.clearBtn}
            onClick={() => {
              cancelSearch();
              setResults([]);
              setSearching(false);
              onChange('', '');
              setExternalFlag(null);
              inputRef.current?.focus();
            }}
            aria-label={strings.Common_PeoplePicker_ClearAriaLabel}
            title={strings.Common_PeoplePicker_ClearAriaLabel}
          >
            ✕
          </button>
        )}
      </div>

      <div
        id={listId}
        role="listbox"
        aria-label={label || strings.Common_PeoplePicker_ListAriaLabel}
        className={styles.dropdown}
        style={listVisible || (open && searching) ? undefined : { display: 'none' }}
      >
        {open && searching && <div className={styles.status} role="presentation">{strings.Common_PeoplePicker_Searching}</div>}
        {directoryOptions.length > 0 && (
          <div role="presentation" className={styles.groupHeader}>{strings.Common_PeoplePicker_DirectoryHeader}</div>
        )}
        {directoryOptions.map(renderOption)}
        {recentOptions.length > 0 && (
          <div role="presentation" className={styles.groupHeader}>{strings.Common_PeoplePicker_RecentHeader}</div>
        )}
        {recentOptions.map(renderOption)}
        {externalOptions.map(renderOption)}
      </div>

      {/* Announces result counts to screen readers as the list changes */}
      <div aria-live="polite" style={VISUALLY_HIDDEN}>
        {open ? (searching
          ? strings.Common_PeoplePicker_Searching
          : formatString(strings.Common_PeoplePicker_ResultsAnnouncement, { count: options.length })) : ''}
      </div>

      {errorMessage && <div id={errorId} className={styles.error} role="alert">{errorMessage}</div>}

      {!!query && !isExternal && email && <div className={styles.hint}>{email}</div>}

      {showEmailField && (
        <div className={styles.emailBlock}>
          <TextField
            label={strings.Common_PeoplePicker_EmailLabel}
            type="email"
            value={email}
            onChange={(_, v) => onChange(name, v || '')}
            placeholder={strings.Common_PeoplePicker_EmailPlaceholder}
            errorMessage={emailInvalid ? strings.Common_PeoplePicker_InvalidEmail : undefined}
          />
          <div className={styles.hint}>{strings.Common_PeoplePicker_ExternalHint}</div>
        </div>
      )}
    </div>
  );
};

export default PeoplePicker;
