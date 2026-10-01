# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

---

## [1.5.0] - 2026-10-01

### Added

- **Copy import errors** — the error and warning lists on the import results screen are now selectable and have a **Copy details** button
- **People picker for Assigned To and Project Manager** — search your directory as you type, pick from recently used people, or enter a free-text name (with an optional email) for external people who aren't in the directory
- **Dependency types and lag** — links can now be Finish to Start (FS), Start to Start (SS), Finish to Finish (FF), or Start to Finish (SF), each with a lag or lead in working days. Set them in the task panel's Links tab; they're stored in the `Dependencies` column in a compact form such as `12,15SS+2` (plain IDs still mean FS with no lag)
- **Auto-shift dependents** — moving or resizing a task moves the tasks that depend on it to keep the link satisfied. A message reports how many dependents moved, with an **Undo** button
- **Undo / Redo** — toolbar buttons and Ctrl+Z / Ctrl+Y for task edits, drags, and dependent shifts (last 50 changes)
- **Working calendar** — new web part property pane group to choose which weekdays are working days and to list holidays (one `yyyy-MM-dd` date per line). Scheduling, keyboard moves, and dependency lag use it, and non-working days are marked in List view
- **Default view and default zoom** — new property pane settings for the view and Gantt zoom a viewer sees first (a viewer's own remembered choice or a shared link still takes priority)
- **Baselines** — **Set Baseline…** saves every task's current dates; the Gantt then shows baseline bars, and the tooltip and task panel show the finish variance versus baseline. The `BaselineStart` and `BaselineDue` columns are added to existing task lists automatically when you have permission to edit the list
- **Gantt interactions**:
  - Live date label while dragging, with auto-scroll when you drag near the edge of the timeline
  - Drag across an empty row to create a task with those dates
  - Drag from a bar's link handle onto another task to create a dependency (circular and duplicate links are rejected)
  - Move and resize bars from the keyboard: focus a bar, then Left/Right arrows move it one working day; Shift + Left/Right changes the end date; Enter opens the task
  - **Collapse all / Expand all phases** buttons; collapsed phases are remembered per project
  - **Fit** button to zoom to the whole project, and Ctrl + mouse wheel to zoom
  - Critical path highlight and critical path list on the Dashboard
- **Copy Link** — copies a link to the current view that includes the project, view, zoom level, and active filters
- **Print / Save as PDF** — from the ⋯ menu
- **CSV and calendar exports** — Export Tasks to CSV (⋯ menu and List view), Portfolio summary CSV, and Export Milestones as an iCal (`.ics`) file
- **Kanban board**:
  - Work-in-progress (WIP) limits per column, with a warning when a column is over its limit
  - Swimlanes by Phase or Assignee
  - Collapsible Canceled column
  - **Add Task** in a column pre-sets that column's status
- **List view** — checkbox selection with bulk set status, priority, % complete, assignee, and bulk delete
- **Dashboard** — Burndown chart, Workload by Assignee (open and overdue counts), and a Critical Path list
- **Portfolio view** — search box, **Hide completed / canceled** toggle, and a remembered sort order
- **Import** — date-order picker (Auto-detect, Month/Day/Year, or Day/Month/Year), required-field markers in the column mapper, and a row validation preview that lists rows with problems (missing name, invalid or reversed dates) with an option to skip them
- **Theme and accessibility** — colors follow the SharePoint site theme and high-contrast mode; added roles, labels, keyboard access, and screen-reader announcements across the Gantt, Kanban, List, Dashboard, Portfolio, and import screens
- **Performance** — the Gantt windows its rows (only visible rows are rendered) and memoizes bars, so large projects scroll and drag more smoothly
- All new text is translated into all 30 supported languages

### Changed

- **"Cancelled" now reads "Canceled"** in the English interface (Status, Kanban column, Dashboard, and project status). The stored value is unchanged, so existing data, filters, and exports keep working
- Dependency parsing on import understands MS Project-style predecessors such as `3FS+2d`
- **Dashboard and phase progress** now count leaf tasks only (parents with sub-tasks are excluded) and ignore canceled tasks; the overdue list is sorted oldest first
- **Portfolio "Status" sort** now follows lifecycle order (Planning, Active, On Hold, Completed, Canceled) instead of alphabetical
- Several hard-coded English strings are now localized
- Dates are stored at 10:00 UTC so a task's calendar day is the same in every time zone (existing UTC-midnight dates are still read correctly)

### Fixed

- **New-project column creation** — a column that SharePoint rejects is now named in the error message; the Dependencies column falls back to a multi-line text column if single-line text rejects its length, ParentTaskId falls back to a non-indexed column, and the optional baseline columns no longer block project creation
- **Concurrent-edit protection now works for tasks** — task saves previously always sent `If-Match: *`, so the 1.3.0 conflict check never fired. Saves now carry the real concurrency token and are serialized per task; on a conflict you see a message and the latest tasks are reloaded
- Throttled requests (HTTP 429, 503, and 504) are now retried with backoff instead of failing
- Filter chips are no longer undone by a search that was still pending, and filters reset whenever you change project
- The Gantt no longer jumps after dragging a bar past the left edge
- Dragging only starts with the primary mouse button, and **Esc** cancels a drag
- Tasks in a circular parent chain now appear at the top level with a warning instead of disappearing
- Kanban: pressing Enter or the arrow keys on a card's Edit or Delete button no longer also triggers the card's own actions
- Task panel: changing status or dependencies now marks the form as changed, so it prompts before discarding
- Import: dependency warnings are now shown; numeric dependencies map by the row's original spreadsheet position; and dates no longer come out one day early in UTC+13/+14 time zones
- Portfolio figures refresh after task and project edits
- Portfolio Excel export includes the header row when there are no projects, and the Portfolio PowerPoint table now paginates across slides
- Exports validate colors, so a malformed color no longer breaks the file
- Deleting a project removes its task list before its registry entry, so a failure can't leave an orphaned entry
- Project creation rolls back if a column fails to be created
- The archived-field check no longer swallows unrelated errors

---

## [1.4.0] - 2026-09-28

### Added

- **Localization** — the web part's UI text is now localized into 30 languages (English plus Arabic, Czech, Danish, German, Greek, Spanish, Finnish, French, Hebrew, Hindi, Hungarian, Indonesian, Italian, Japanese, Korean, Norwegian Bokmål, Dutch, Polish, Portuguese (Brazil and Portugal), Romanian, Russian, Swedish, Thai, Turkish, Ukrainian, Vietnamese, and Chinese (Simplified and Traditional)), matching the current SharePoint UI language automatically

### Fixed

- **Project selector in Portfolio view** — selecting a different project from the toolbar's project dropdown while viewing Portfolio now switches to that project's Gantt view instead of appearing to do nothing (the view mode wasn't leaving Portfolio when a project was picked)

---

## [1.3.0] - 2026-07-14

### Added

- **Concurrent-edit protection** — task and project saves now carry a SharePoint concurrency token (etag). If someone else saved the same item since you loaded it, your save is rejected with *"This task/project was changed by someone else since you loaded it. Refresh and try again."* instead of silently overwriting their change
- **Touch and pen support on the Gantt chart** — dragging to move a bar and dragging the edge to resize now work with touch and pen input, not just a mouse (Gantt bars use Pointer Events internally)
- **Keyboard accessibility**:
  - List view column headers can be sorted with Enter/Space, not just a click
  - Kanban cards can be opened with Enter and moved between status columns with the Left/Right arrow keys
  - The toolbar's project selector and every "⋯" context menu are fully keyboard-operable (Tab + Enter/Space)
  - Gantt display-settings toggles and swatches now have proper accessible names and keyboard focus
- **Portfolio "Stats unavailable" state** — a project whose task stats fail to load (permissions, throttling, network) now shows as explicitly unavailable in the Portfolio view and in both portfolio exports, instead of silently rendering as a healthy, empty project
- **Positional dependency resolution on import** — predecessors referenced by row number or by title are now matched against the tasks actually created from that import, with a warning surfaced (rather than a silent wrong match) when a referenced title is ambiguous

### Changed

- **Import — percent-complete columns**: a column stored as a true Excel percentage (displaying "50%" but holding the raw value `0.5`) now imports as 50, not 0.5
- **Import — appending to an existing project**: newly imported tasks are now sorted after the project's existing tasks instead of interleaving with them
- **Import — Planner source**: now pages through all of a user's groups and all of a plan's tasks instead of stopping at the first page, so large plans/tenants no longer import partially with no warning
- **Import — task creation**: batched into chunked REST requests instead of one request per task, substantially speeding up large imports
- **Gantt "Critical path highlight"** now only shows the red highlight when the toggle is actually on, instead of whenever critical-path data was being computed for the arrow-display settings
- **Gantt "Bar Style" (Flat/Gradient)** now applies to the on-screen chart, matching what was already true of the exported image
- **List view sorting** — Priority and Status columns now sort in their natural order (Critical → Low; workflow order) instead of alphabetically; tasks with no date always sort last regardless of sort direction
- **Dashboard "Updated this week" / "Completed this week"** now show the most recently modified tasks instead of an arbitrary subset
- **Deleting a task with sub-tasks** now pages through all of its children (previously capped at 500) and aborts the delete instead of proceeding if re-parenting any child fails

### Fixed

- Gantt timeline no longer snaps back to the earliest task after every task edit or drag — it now only re-centers when the visible date range genuinely changes
- A milestone with only a due date set (no start date) now renders at the correct date instead of snapping to today, and its dependency arrow no longer detaches from the diamond
- A sub-task of a sub-task ("grandchild" task) no longer silently disappears from the Gantt and List views; the task panel also now blocks nesting a task under another task if it already has sub-tasks of its own
- Fixed a race where two people creating projects with the same name at nearly the same time could hit a raw duplicate-name error instead of one of them getting a usable fallback name
- Fixed a SharePoint list being left behind with no registry entry if project field provisioning failed partway through project creation
- Fixed the Kanban board's drop-target highlight flickering while dragging a card across other cards in the same column
- Fixed the toolbar search box causing the entire visible view (Gantt/List/Kanban) to re-render on every keystroke — input is now debounced
- Fixed task/project save errors being shown behind the still-open edit panel instead of inside it, which could look like the save had silently failed (or succeeded) either way
- Fixed a numeric dependency reference from Excel/MS Project imports potentially wiring a task's predecessor to the wrong task when titles repeated in the source file
- Fixed switching the import source between Excel and Planner leaving stale data from whichever source was previously selected
- Fixed rapid-clicking between two Planner plans in the import panel potentially importing the wrong plan's tasks
- Fixed the "Dependency arrows" Gantt display setting always resetting to on after a page reload, ignoring the saved preference
- Fixed a task update potentially applying to the wrong project's task list if the user switched projects while a save was still in flight
- Fixed the custom color picker (task/project panels) producing an invalid color when given a malformed hex value
- Fixed the Web Part Title property pane setting having no visible effect — it now renders above the toolbar when set
- Corrected two stale doc claims: Display Settings persist across reloads via browser local storage (they were never truly session-only), and the bundled `xlsx` package is pinned to SheetJS 0.20.3, not 0.18.5

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
