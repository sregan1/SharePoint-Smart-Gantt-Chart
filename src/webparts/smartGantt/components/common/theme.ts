// Theme-aware colors for chrome (backgrounds, text, borders). SharePoint
// exposes its theme slots as CSS variables on modern pages; the fallbacks are
// the Office default palette so the web part still renders outside SharePoint
// (workbench, tests). Semantic status/priority/health colors are NOT here —
// they stay in models/index.ts and utils/healthUtils.ts.
//
// Only use these where the value is used as-is: they can't take a hex alpha
// suffix (`${color}30`) like a plain #RRGGBB string can.
export const THEME = {
  text: 'var(--neutralPrimary, #323130)',
  textSecondary: 'var(--neutralSecondary, #605E5C)',
  textTertiary: 'var(--neutralTertiary, #8A8886)',
  border: 'var(--neutralLight, #EDEBE9)',
  borderStrong: 'var(--neutralQuaternary, #C8C6C4)',
  surface: 'var(--white, #fff)',
  surfaceAlt: 'var(--neutralLighterAlt, #FAF9F8)',
  surfaceMuted: 'var(--neutralLighter, #F3F2F1)',
  accent: 'var(--themePrimary, #0078D4)',
  accentSurface: 'var(--themeLighterAlt, #EFF6FC)',
} as const;

// Shared by the inline-styled views (Dashboard, Portfolio), which can't use a
// CSS module for pseudo-classes and media queries. Class names are
// namespaced with "sg" to avoid clashing with host-page styles.
export const SHARED_VIEW_CSS = `
  .sgFocusable:focus-visible { outline: 2px solid var(--themePrimary, #0078D4); outline-offset: 2px; }
  @media (forced-colors: active) {
    .sgCard { border: 1px solid CanvasText !important; }
    .sgFill { forced-color-adjust: none; }
    .sgFocusable:focus-visible { outline-color: Highlight; }
  }
`;
