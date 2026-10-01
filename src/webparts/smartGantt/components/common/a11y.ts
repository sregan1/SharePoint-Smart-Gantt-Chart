import * as React from 'react';

/**
 * Builds a keydown handler that activates a non-<button> "button" on Enter or
 * Space. Ignores events bubbling up from focusable descendants (inner
 * buttons, inputs) so they don't trigger the outer element's action too.
 */
export function onActivate(action: () => void): (e: React.KeyboardEvent) => void {
  return (e: React.KeyboardEvent): void => {
    if (e.target !== e.currentTarget) return;
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      action();
    }
  };
}

/** Inline style for text that is announced by screen readers but not shown. */
export const VISUALLY_HIDDEN: React.CSSProperties = {
  position: 'absolute',
  width: 1,
  height: 1,
  margin: -1,
  padding: 0,
  overflow: 'hidden',
  clip: 'rect(0 0 0 0)',
  whiteSpace: 'nowrap',
  border: 0,
};

let idCounter = 0;
/** Stable unique DOM id for the lifetime of a component instance (React 17 has no useId). */
export function useUniqueId(prefix: string): string {
  const ref = React.useRef<string>();
  if (!ref.current) ref.current = `${prefix}-${++idCounter}`;
  return ref.current;
}
