# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

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

[Unreleased]: https://github.com/sharepointsmartsolutions/SharePointSmartGanttChart/compare/v1.1.0...HEAD
[1.1.0]: https://github.com/sharepointsmartsolutions/SharePointSmartGanttChart/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/sharepointsmartsolutions/SharePointSmartGanttChart/releases/tag/v1.0.0
