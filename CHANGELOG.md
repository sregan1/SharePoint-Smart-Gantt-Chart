# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

---

## [1.2.1] - 2026-06-17

### Added

- **Import File as New Project** — new entry in the project selector dropdown (📥 Import File as New Project…) that opens the import panel in project-creation mode; creates a new project automatically from an Excel or CSV file without requiring a project to exist first
- **Predecessors column in List View** — shows the names of predecessor tasks each task depends on; displays "—" for tasks with no dependencies; falls back to `#id` if a referenced task has been deleted
- **Dependency arrow visibility controls** — two new sub-options under the "Dependency arrows" Show/Hide toggle (only visible when Dependency arrows is on):
  - **Critical path always visible** — critical path arrows are permanently drawn regardless of hover state; critical path is now computed whenever this option is on, even if "Critical path highlight" is off
  - **All others on hover only** — non-critical dependency arrows only appear when hovering a task bar; hovering always reveals all dependency arrows for that task regardless of the critical-path filter; both options default to on for cleaner views on dense MS Project imports

### Changed

- **Dependency arrow routing** — all connector lines now use orthogonal (90°) paths exclusively for a clean, consistent layout matching professional Gantt tools:
  - **Forward dependencies** — replaced smooth S-curves (cubic Bézier) with a right-angle elbow: right to the midpoint between the two bars, drop straight down to the target row, then right into the bar
  - **Backward/overlapping dependencies** — routing lane sits in the mid-gap between adjacent rows (50% of row height above/below target center), never overlapping the bar; corrected routing to approach from outside the row rather than passing through it
  - **Milestone dependencies** — arrow exits from the bottom tip of the diamond, drops straight down to the target row, then travels right into the bar (two segments max); same-date or backward dependencies drop straight into the top of the bar with a downward arrowhead
  - **Arrowhead** — reduced in size for a less dominant appearance at all zoom levels
- **Compact row height** — increased from 32 → 36 px, giving dependency routing lanes 5 px clearance from each adjacent bar (vs. 3 px previously)

### Fixed

- **Gantt / left-panel row alignment** — added `box-sizing: border-box` to `.taskRow` and `.leftHeader` so the 1 px `border-bottom` is counted within the declared height, eliminating progressive vertical drift between the task-name list and the SVG timeline that accumulated over many rows
- **Milestone arrow origin** — arrow now uses `startDate` (matching where `renderBar` draws the diamond) instead of `dueDate`; fixes the arrow appearing displaced to the right when a milestone's start and due dates differ
- **Web part component ID** — updated to a new GUID to allow clean re-deployment when a conflicting prior version exists in the SharePoint App Catalog

---

## [1.2.0] - 2026-06-13

### Added

- **Task Filter Bar** — a new third toolbar row that appears when a project has tasks. Six controls let you narrow the visible task list in real time:
  - **Search** — text input that matches against task names (case-insensitive)
  - **Status** — multi-select chip dropdown (Not Started · In Progress · Completed · On Hold · Cancelled)
  - **Priority** — multi-select chip dropdown (Critical · High · Medium · Low)
  - **Assignee** — multi-select chip dropdown; appears only when the project has assigned tasks
  - **Phase** — multi-select chip dropdown; appears only when the project has phases
  - **Due date** — single-select dropdown: *Any due date*, *Overdue*, *Due today*, *Due in 7 days*
  - Active filter chips are highlighted in blue with a count badge; a **match count** (e.g. "5 of 20") appears when any filter is active; a **✕ Clear filters** button resets everything at once
  - Filter state persists when switching between Gantt, List, Kanban, and Dashboard views; exports always use the full unfiltered task list
- **`FilterBar` component** (`src/webparts/smartGantt/components/toolbar/FilterBar.tsx`) — self-contained filter UI using Fluent UI `Callout` and `Checkbox`; `MultiChip` sub-component for multi-select dropdowns
- **`filterUtils.ts`** (`src/webparts/smartGantt/utils/filterUtils.ts`) — pure utility module: `filterTasks(tasks, filter)` and `isFilterActive(filter)`; no React dependency
- **`dateUtils.ts`** (`src/webparts/smartGantt/utils/dateUtils.ts`) — pure date utility functions extracted from component code for reuse across views
- **`ITaskFilter` interface** and **`DueFilter` type** added to `models/index.ts`; `EMPTY_TASK_FILTER` constant for safe initialization; `isFilterActive` re-exported from models

### Changed

- **Toolbar** — adds a Row 3 (`styles.row3`) that renders `FilterBar` when `viewMode !== 'portfolio'`, a project is selected, and `totalCount > 0`; `filteredCount` and `totalCount` props added
- **`SmartGantt.tsx`** — `taskFilter` and `portfolioStats` state added; `filterTasks()` applied before rendering views so all four views (Gantt, List, Kanban, Dashboard) receive the filtered task set
- **GanttChart** — rendering and interaction improvements
- **ListView** — updated to consume filtered tasks passed from root; column and layout refinements
- **KanbanView** — updated to consume filtered tasks; card and column improvements
- **DashboardView** — updated to consume filtered tasks; stats recalculated from filtered set
- **SharePointService** — query and field-selection improvements across project and task operations
- **ImportService** — column mapping and import reliability improvements

---

## [1.1.0] - 2026-06-08

### Added

- **Portfolio View** — cross-project overview accessible from the project selector dropdown (⊞ Portfolio). Shows all projects as summary cards with colored left border, manager avatar, status badge, computed health badge, overall progress bar, task count breakdown (Done / Active / At Risk / Overdue), and a mini date-range timeline with a today marker. Sortable by name, health, status, or completion %.
- **Health Status Indicators** — automatic On Track / At Risk / Overdue / Done badges computed at render time from each task's dates and `percentComplete`. Never stored; derived from the existing task data. Logic: `Completed/Cancelled → Done`, past due date → `Overdue`, start date passed but Not Started → `At Risk`, progress more than 10 % behind schedule → `At Risk`, otherwise → `On Track`.
- **Project health rollup** — portfolio cards show the worst health across Critical/High priority tasks. Overdue on a low-priority task escalates to At Risk at the project level.
- **Health column in List view** — new non-sortable Health column between Status and Priority.
- **Health badges on Kanban cards** — badge shown in the card tag row alongside status and phase.
- **Health badge in Gantt tooltip** — shown when hovering a task bar (respects the Show/Hide toggle).
- **"By Health" color coding** — new option in Display Settings Color Coding section; Gantt bars recolor to match task health (blue = On Track, orange = At Risk, red = Overdue, green = Done).
- **"Health status badges" toggle** — new Show/Hide option in Display Settings; when off, health badges are hidden across List, Kanban, and Gantt tooltip.
- **Portfolio export to Excel** — downloads `Portfolio Summary.xlsx` with one row per project: name, status, health, task counts, % done, start, and due date.
- **Portfolio export to PowerPoint** — downloads `Portfolio Report.pptx` with a cover slide (aggregate health summary) and a color-coded project summary table slide.
- **Archive project** — projects can be archived (hidden from the project selector and portfolio by default). A "Show archived projects" toggle appears in the project selector dropdown whenever at least one archived project exists. Archived projects display at reduced opacity with an "Archived" badge.
- **Unarchive project** — when an archived project is selected (via "Show archived"), the ⋯ menu shows "Unarchive Project" to restore it.
- **Automatic field migration** — `ensureProjectsList()` detects and adds the new `IsArchived` boolean field to existing project registries on first load after upgrade; no manual schema update required.
- **`HealthBadge` component** (`src/webparts/smartGantt/components/common/HealthBadge.tsx`) — reusable pill badge with `sm` (dot + label) and `md` (padded pill) sizes.
- **`healthUtils.ts`** (`src/webparts/smartGantt/utils/healthUtils.ts`) — pure utility functions: `computeTaskHealth`, `computeProjectHealth`, `healthColor`, `healthLightColor`, `healthLabel`. No React dependency.
- **`IProjectTaskStats` interface** — lightweight per-project aggregate loaded with a minimal SharePoint field select (Status, Priority, PercentComplete, StartDate, DueDate, IsMilestone).
- **`getProjectTaskStats` / `getAllProjectStats`** in `SharePointService.ts` — stats are fetched in parallel with `Promise.all` on first portfolio visit and cached until the next task or project mutation.
- **`PortfolioView` component** (`src/webparts/smartGantt/components/views/PortfolioView.tsx`) — full portfolio card grid with header bar aggregate health summary, sort controls, and refresh button.
- **`exportPortfolioToExcel` / `exportPortfolioToPowerPoint`** in `ExportService.ts`.
- **MIT `LICENSE` file** added to repository root.
- **`docs/sample-data/Tech-Conference-Tasks.xlsx`** updated — conference dates shifted to October 2026.

### Changed

- **Delete project** now moves the project and its task list to the **SharePoint recycle bin** (`.recycle()`) instead of permanently deleting them. Items are recoverable from the recycle bin for up to 93 days. Confirmation dialog retitled "Send to Recycle Bin?".
- **Project selector dropdown** sorts projects alphabetically (was: by creation date).
- Project selector dropdown now includes a **Portfolio** entry at the top, above the project list.
- Toolbar row 2 (zoom controls, view switcher, ⋯ menu) is hidden when in portfolio mode; a portfolio-specific ⋯ menu with export options appears in row 1 instead.
- Portfolio stats are loaded **lazily** on first navigation to Portfolio and invalidated on any task or project create, edit, or delete.
- Display Settings Color Coding section now has four options: By Status, By Priority, By Phase, **By Health**.
- Display Settings Show/Hide section now has five toggles, adding **Health status badges**.
- `USER_GUIDE.md` renamed to `USER-GUIDE.md`; badge link in `README.md` updated accordingly.
- Screenshots updated: `screenshot-list.png` (Health column), `screenshot-kanban.png` (health badges), `screenshot-display-settings.png` (By Health + health badge toggle), `screenshot-portfolio.png` (new).

---

## [1.0.0] - 2026-06-04

### Added

- Initial public release.
- **Gantt Chart** — custom SVG timeline with drag-to-move, drag-to-resize, dependency arrows, phase grouping, four zoom levels (Day / Week / Month / Quarter), and a today indicator.
- **List View** — sortable grid with inline status and priority editing, overdue highlighting, and progress bars.
- **Kanban Board** — five-column drag-and-drop board (Not Started → In Progress → On Hold → Completed → Cancelled).
- **Dashboard View** — summary stats and recent activity feed.
- **Task Panel** — three-tab side panel (Basic, Details, Links) for creating and editing tasks; supports subtask hierarchy and dependency linking.
- **Project management** — each project gets its own SharePoint list with 15 pre-built columns; project color, status, dates, and description are configurable.
- **Display Settings** — color coding (By Status / Priority / Phase), header theme, week numbering (ISO or project-relative), bar style (Gradient / Flat), row height (Compact / Normal / Spacious), and five Show/Hide toggles.
- **Export to Excel** — downloads all task columns as a formatted `.xlsx` file.
- **Export to PowerPoint** — four-slide deck: cover, project summary, Gantt timeline, and recent activity.
- **Export as Image (PNG)** — full-resolution Gantt chart rendered as a 2× PNG.
- **Import from Excel/CSV** — column-mapping screen with auto-detection of common headers; supports MS Project exports.
- **Import from Microsoft Planner** — reads Planner plans via Microsoft Graph; buckets become phases.
- **Autocomplete** — Phase and Assigned To fields suggest values already used in the project.
- **Guest user support** — Microsoft 365 B2B guests with Site Member permission can view and edit tasks in existing projects.
- SPFx 1.20.0 · React 17 · Fluent UI 8 · PnPjs 3 · SheetJS · date-fns · pptxgenjs.

[Unreleased]: https://github.com/sharepointsmartsolutions/SharePointSmartGanttChart/compare/v1.2.1...HEAD
[1.2.1]: https://github.com/sharepointsmartsolutions/SharePointSmartGanttChart/compare/v1.2.0...v1.2.1
[1.2.0]: https://github.com/sharepointsmartsolutions/SharePointSmartGanttChart/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/sharepointsmartsolutions/SharePointSmartGanttChart/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/sharepointsmartsolutions/SharePointSmartGanttChart/releases/tag/v1.0.0
