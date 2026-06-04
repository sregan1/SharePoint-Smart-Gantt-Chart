# SharePoint Smart Gantt Chart

[![Website](https://img.shields.io/badge/Website-sharepointsmartsolutions.com-0078D4?style=for-the-badge&logo=microsoftsharepoint&logoColor=white)](http://sharepointsmartsolutions.com/smart-gantt)
[![User Guide](https://img.shields.io/badge/User_Guide-Read%20the%20Docs-107C10?style=for-the-badge&logo=readthedocs&logoColor=white)](USER_GUIDE.md)

A SharePoint Framework (SPFx) web part for project management with three views — Gantt chart, list, and Kanban board — all backed by SharePoint lists.

![SPFx](https://img.shields.io/badge/SPFx-1.20.0-blue) ![Node](https://img.shields.io/badge/Node-18.x-green) ![React](https://img.shields.io/badge/React-17-blue) [![License: MIT](https://img.shields.io/badge/license-MIT-lightgrey)](LICENSE)

![Gantt Chart View](docs/screenshots/screenshot-gantt.png)

---

## Features

- **Gantt Chart** — Custom SVG timeline with drag-to-move, drag-to-resize, dependency arrows, phase grouping, zoom levels (Day / Week / Month / Quarter), and a today indicator
- **List View** — Sortable, Excel-style grid with inline status and priority editing, overdue highlighting, and progress bars
- **Kanban Board** — Drag-and-drop cards across status columns (Not Started → In Progress → On Hold → Completed → Cancelled)
- **Project management** — Each project gets its own SharePoint list with 15 pre-built columns (status, priority, dates, assignee, % complete, phase, milestones, dependencies, and more)
- **Display settings** — Customize colors, header theme, week numbering, bar style, row height, and more for the Gantt view
- **Export** — Download tasks as Excel, export a full PowerPoint project report (cover, summary, Gantt chart, and recent activity), or save the Gantt as a high-resolution PNG
- **Import** — Bring in tasks from Excel/CSV files (including MS Project exports) or directly from Microsoft Planner, with a column-mapping screen for non-standard headers
- **Autocomplete** — Phase and Assigned To fields suggest values already used in the project

---

## Views

### Gantt Chart

![Gantt Chart View](docs/screenshots/screenshot-gantt.png)

- Colored project title bar at the top of the timeline
- Two-row toolbar: project/task actions on the first row; view and zoom controls on the second
- Sticky task list on the left; scrollable SVG timeline on the right
- Four zoom levels: Day, Week, Month, Quarter
- Task bars color-coded by status, priority, or phase; progress overlay shows % complete
- Drag a bar horizontally to move dates; drag the right edge to resize
- Dependency arrows drawn between tasks
- Phase rows collapse/expand to group related tasks
- Hover tooltip shows task name, dates, status, priority, assignee, and % complete
- Today line with red indicator

### List View

![List View](docs/screenshots/screenshot-list.png)

- Click any column header to sort ascending/descending
- Change status or priority inline via dropdown — saves to SharePoint immediately
- Overdue tasks highlighted in red
- Phase group rows visually separate tasks
- Subtasks indented under their parent

### Kanban Board

![Kanban View](docs/screenshots/screenshot-kanban.png)

- Five columns matching task statuses
- Drag cards between columns to update status
- Cards show priority color, tags, due date, assignee avatar, and progress bar
- Add Task button in each column

### Task Panel

Click any task name or **+ Add Task** to open the task panel. It has three tabs:

| Tab | Key fields |
|---|---|
| **Basic** | Name, description, dates, status, priority, % complete slider, assigned to |
| **Details** | Phase (auto-colors the bar), milestone toggle, custom bar color picker, notes |
| **Links** | Parent task (sub-task hierarchy), dependencies (removable chips + dropdown) |

![Task panel — Details tab](docs/screenshots/screenshot-task-panel-details.png)
![Task panel — Links tab](docs/screenshots/screenshot-task-panel-links.png)

---

## Display Settings

Click **⚙ Display** in the toolbar while in Gantt view to open the settings panel.

| Setting | Options |
|---|---|
| **Color Coding** | By Status, By Priority, By Phase (each phase gets a consistent auto-color) |
| **Header Color** | Dark (default), Navy, Teal, Purple, Light |
| **Week Numbering** | ISO weeks (W23, W24…) or Project-relative (W1, W2, W3… from the first task's start date) |
| **Bar Style** | Gradient or Flat |
| **Row Height** | Compact (32px), Normal (40px), Spacious (52px) |
| **Show / Hide** | Weekend shading, dependency arrows, progress % on bars, assignee name on bars |

Settings are applied live and remembered for the session.

**Project-relative week numbers** are particularly useful for presentations — stakeholders can refer to "Week 3" without needing to know the calendar date.

---

## Exporting

![Export menu](docs/screenshots/screenshot-export-menu.png)

All export options are in the **⋯ menu** (top-right of the toolbar).

### Export to Excel

Downloads `<ProjectName> - Tasks.xlsx` with all task columns (name, phase, dates, status, priority, assignee, % complete, milestone flag, notes). Columns are auto-sized to their content.

### Export to PowerPoint

![PowerPoint export — Cover, Summary, and Gantt slides](docs/screenshots/screenshot-pptx-export.png)

Downloads `<ProjectName> - Project Report.pptx` — a four-slide deck:

| Slide | Contents |
|---|---|
| **Cover** | Project title, status, date range, description, and project manager |
| **Summary** | Task counts by status, overall progress bar, and status breakdown |
| **Gantt Timeline** | Full Gantt chart rendered as a high-resolution PNG image |
| **Summary & Recent Activity** | Project overview on the left; tasks completed or updated in the past 7 days on the right |

### Export as Image (PNG)

Renders the full Gantt chart — every task, the complete date range, the project title bar, and the current color/theme settings — as a clean 2× high-resolution PNG. No browser window cropping. Suitable for pasting directly into Word or email.

The export uses the current Display Settings, so you can tune colors and layout before exporting.

---

## Requirements

| Requirement | Version |
|---|---|
| Node.js | 18.x |
| SharePoint Online | Any |
| SPFx | 1.20.0 |
| Permissions | Site Member or above (list creation requires Site Owner on first use) |

For **Planner import**, a Microsoft 365 admin must approve Graph API permissions once after deployment (see [Planner Import Setup](#planner-import-setup)).

---

## External / Guest User Access

The web part supports **Microsoft 365 guest users** (Azure AD B2B external collaborators) with a small set of caveats.

### What guests can do

Once a guest has been invited to the SharePoint site and granted **Site Member (Contribute)** permission by a Site Owner, they can:

| Feature | Supported |
|---|---|
| View projects, tasks, and all views (Gantt, List, Kanban, Dashboard) | ✅ |
| Add, edit, and delete tasks | ✅ |
| Export to Excel, PowerPoint, and PNG | ✅ (all client-side) |
| Be assigned to tasks | ✅ (Assigned To is a plain-text field — no Azure AD lookup required) |
| Import tasks from Excel / CSV | ✅ |

### What guests cannot do

| Feature | Limitation | Workaround |
|---|---|---|
| **Create new projects** | Requires "Manage Lists" permission — guests cannot be Site Owners | An internal Site Owner or Member creates the project; guests work within it |
| **Planner import** | Requires `Group.Read.All` Graph permission; guests typically cannot read organizational Planner plans | Export the Planner plan to Excel first, then import via the Excel path |

### Setting up guest access

1. In the SharePoint site, go to **Settings → Site Permissions → Invite People**
2. Enter the guest's email address — they'll receive an invitation email
3. Grant them **Edit** (Contribute) permission, which maps to Site Member access
4. An internal Site Owner should create any new projects before the guest arrives; the guest can then manage all tasks within those projects

> **Note:** Guests must accept the SharePoint invitation and sign in with their Microsoft account (personal, work, or school) before they can access the site. This is standard Microsoft 365 B2B guest behavior and is not specific to this web part.

---

## Getting Started

### 1. Clone and install

```bash
git clone <repo-url>
cd SharePointSmartGanttChart
npm install
```

### 2. Configure the workbench URL

`config/serve.json` is excluded from the repo (it contains your tenant URL). Copy the example file and fill in your tenant:

```bash
cp config/serve.json.example config/serve.json
```

Then edit `config/serve.json`:

```json
{
  "port": 4321,
  "https": true,
  "initialPage": "https://<your-tenant>.sharepoint.com/_layouts/15/workbench.aspx"
}
```

### 3. Trust the dev certificate (first time only)

```bash
gulp trust-dev-cert
```

### 4. Run the dev server

```bash
gulp serve
```

The browser will open to your SharePoint workbench. Add the **Smart Gantt Chart** web part from the toolbox.

---

## Building for Production

```bash
# Bundle (minified)
gulp bundle --ship

# Package the .sppkg
gulp package-solution --ship
```

Upload `sharepoint/solution/sharepoint-smart-gantt-chart.sppkg` to your **SharePoint App Catalog**.

---

## How It Works

### SharePoint lists

The web part creates and manages two types of lists on the current site:

| List | Purpose |
|---|---|
| `SmartGantt_Projects` | Registry of all projects (created automatically on first use) |
| `SGP_<ProjectName>_<id>` | One list per project, containing all tasks |

**Project list columns** (created automatically):

| Column | Type | Notes |
|---|---|---|
| Title | Text | Task name |
| TaskDescription | Note | Optional description |
| StartDate | DateTime | |
| DueDate | DateTime | |
| Status | Choice | Not Started / In Progress / Completed / On Hold / Cancelled |
| Priority | Choice | Critical / High / Medium / Low |
| PercentComplete | Number | 0–100 |
| AssignedToName | Text | Assignee display name |
| AssignedToEmail | Text | Assignee email |
| Phase | Text | Groups tasks on the Gantt |
| IsMilestone | Boolean | Renders as a diamond on the Gantt |
| ParentTaskId | Number | ID of parent task (for subtask hierarchy) |
| Dependencies | Text | Comma-separated task IDs |
| Notes | Note | Rich notes field |
| TaskColor | Text | Hex color override (auto-colors by status if blank) |
| SortOrder | Number | Display order |

---

## Importing Tasks

Access import from the **⋯ menu** → **Import Tasks…**

### From Excel or CSV

1. Drop or browse for a `.xlsx`, `.xls`, `.csv`, or `.ods` file
2. The importer reads the first sheet and detects headers
3. Common column names are auto-mapped (e.g. "Owner" → Assigned To, "Finish" → Due Date)
4. Unrecognized columns appear in the column mapper — assign each to a Smart Gantt field or mark as Skip
5. A preview shows the first 3 rows with mapped values
6. Click **Import** to create all tasks

**Microsoft Project Desktop:** Use File → Save As → Excel Workbook (.xlsx) in Project, then import that file. All standard Project columns are recognized automatically.

### From Microsoft Planner

1. Select **Microsoft Planner** as the source
2. Browse the list of Planner plans you have access to
3. Select a plan — tasks load automatically
4. Planner buckets become Phases; assignments, dates, and % complete are mapped automatically
5. Click **Import**

#### Planner Import Setup

Planner import requires Graph API permissions approved once by a Microsoft 365 admin:

1. Deploy the `.sppkg` to the App Catalog
2. In the **SharePoint Admin Center**, go to **Advanced → API Access**
3. Approve the following requests:
   - `Microsoft Graph — Tasks.Read`
   - `Microsoft Graph — Group.Read.All`
   - `Microsoft Graph — User.ReadBasic.All`

These permissions are tenant-wide and only need to be approved once.

---

## Project Structure

```
src/
└── webparts/smartGantt/
    ├── SmartGanttWebPart.ts              # Web part entry point
    ├── SmartGanttWebPart.manifest.json
    ├── models/
    │   └── index.ts                      # IProject, ITask, IGanttDisplaySettings,
    │                                     # color constants, theme definitions
    ├── services/
    │   ├── SharePointService.ts          # All SharePoint list operations (PnPjs)
    │   ├── ImportService.ts              # Excel parsing, Planner Graph calls, batch import
    │   └── ExportService.ts              # Excel export, SVG/PNG Gantt rendering
    └── components/
        ├── SmartGantt.tsx                # Root component — state, routing between views
        ├── toolbar/
        │   └── Toolbar.tsx               # Two-row toolbar: project actions + view controls
        ├── gantt/
        │   ├── GanttChart.tsx            # Custom SVG Gantt chart
        │   └── GanttSettings.tsx         # Display settings panel
        ├── views/
        │   ├── ListView.tsx              # Sortable grid view
        │   └── KanbanView.tsx            # Drag-and-drop Kanban board
        ├── panels/
        │   ├── ProjectPanel.tsx          # Create / edit project side panel
        │   └── TaskPanel.tsx             # Create / edit task side panel (3 tabs)
        ├── common/
        │   └── AutocompleteField.tsx     # Reusable keyboard-navigable suggestion input
        └── import/
            ├── ImportPanel.tsx           # 4-step import flow
            └── ColumnMapper.tsx          # Column mapping UI with auto-map and preview
```

---

## Tech Stack

| Library | Version | Purpose |
|---|---|---|
| SPFx | 1.20.0 | SharePoint web part framework |
| React | 17.0.1 | UI |
| Fluent UI | 8.125.6 | Microsoft design system components |
| PnPjs | 3.26.0 | SharePoint REST API client |
| SheetJS (xlsx) | 0.18.5 | Excel/CSV file parsing and export |
| date-fns | 2.30.0 | Date calculations for Gantt rendering and export |

Microsoft Graph is accessed via the SPFx built-in `msGraphClientFactory` — no extra SDK required.

---

## Configuration

The web part has one property pane setting:

| Property | Default | Description |
|---|---|---|
| Title | Smart Gantt Chart | Web part display title |

All other configuration (projects, tasks, colors, display settings) is managed through the web part UI itself.

---

## Known Limitations

- **MS Project Desktop (.mpp files)** — the binary `.mpp` format cannot be parsed in a browser. Use File → Save As → Excel in Project Desktop instead.
- **Planner import** requires admin approval of Graph permissions (one-time setup).
- **Dependencies** imported from Excel are stored as text and not auto-resolved to task IDs across imports.
- **Display settings** are session-only and reset on page reload. Future work could persist them to the web part property bag.
- The web part requires **Site Owner** permissions on the SharePoint site for the first project creation (list creation). Subsequent task operations work with Site Member permissions.
- **Guest users** cannot create projects (list creation requires elevated permissions), but can fully manage tasks in existing projects. See [External / Guest User Access](#external--guest-user-access) for setup instructions.
