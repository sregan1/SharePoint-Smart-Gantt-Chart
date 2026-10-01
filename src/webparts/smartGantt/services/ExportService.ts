import * as XLSX from 'xlsx';
import PptxGenJS from 'pptxgenjs';
import {
  addDays, differenceInCalendarDays, endOfMonth, format,
  isWeekend, startOfMonth, startOfWeek, addWeeks, addMonths, max, min, getISOWeek,
} from 'date-fns';
import * as strings from 'SmartGanttWebPartStrings';
import {
  IProject, ITask, IProjectTaskStats, IGanttDisplaySettings, STATUS_COLORS, PRIORITY_COLORS,
  HEADER_THEME_COLORS, phaseColor,
} from '../models';
import { computeTaskHealth, healthColor, healthLabel } from '../utils/healthUtils';
import { serializeDependencies } from '../utils/dependencyUtils';
import { parseDateOnly, formatDateOnly, todayLocalMidnight, dateToDateOnlyString as dateOnlyString } from '../utils/dateUtils';
import { formatString } from '../components/localeUtils';

// Strip characters that are invalid in file names (project titles can contain
// anything).
function safeFileName(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, '-').trim();
}

// Colors come from user-editable data (project/task color fields) and end up
// inside SVG markup and PowerPoint XML, so never interpolate them raw.
// Accepts #rgb / #rgba / #rrggbb / #rrggbbaa and CSS named colors; anything
// else (including markup) falls back. Always returns lowercase '#rrggbb' (alpha
// dropped) so callers can append their own alpha suffix.
const HEX_COLOR_RE = /^#([0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})$/i;
const DEFAULT_COLOR = '#0078d4';

function expandHex(c: string): string {
  let h = c.slice(1).toLowerCase();
  if (h.length === 3 || h.length === 4) h = h.split('').map(ch => ch + ch).join('');
  return `#${h.slice(0, 6)}`;
}

let colorProbe: CanvasRenderingContext2D | null | undefined;

function namedColorToHex(name: string): string | null {
  try {
    if (colorProbe === undefined) colorProbe = document.createElement('canvas').getContext('2d');
    if (!colorProbe) return null;
    // The canvas keeps the previous fillStyle when it can't parse the new one,
    // so probe with two different sentinels to tell "invalid" from "same color".
    colorProbe.fillStyle = '#010203';
    colorProbe.fillStyle = name;
    const first = String(colorProbe.fillStyle);
    colorProbe.fillStyle = '#040506';
    colorProbe.fillStyle = name;
    const second = String(colorProbe.fillStyle);
    return first === second && HEX_COLOR_RE.test(first) ? expandHex(first) : null;
  } catch {
    return null;
  }
}

function safeColor(value: string | undefined | null, fallback: string = DEFAULT_COLOR): string {
  const c = (value || '').trim();
  if (HEX_COLOR_RE.test(c)) return expandHex(c);
  if (/^[a-z]{3,20}$/i.test(c)) {
    const named = namedColorToHex(c);
    if (named) return named;
  }
  return HEX_COLOR_RE.test(fallback) ? expandHex(fallback) : DEFAULT_COLOR;
}

function downloadBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  // Some browsers start the download asynchronously; revoke on the next tick.
  setTimeout(() => URL.revokeObjectURL(url), 0);
}

// ─── Excel export ─────────────────────────────────────────────────────────────

const TASK_EXPORT_HEADERS = [
  strings.Export_ColTaskName, strings.Export_ColPhase, strings.Export_ColStartDate, strings.Export_ColDueDate,
  strings.Export_ColStatus, strings.Export_ColPriority, strings.Export_ColAssignedTo, strings.Export_ColAssignedToEmail,
  strings.Export_ColPercentComplete, strings.Export_ColIsMilestone, strings.Export_ColDescription, strings.Export_ColNotes,
];

export function exportTasksToExcel(project: IProject, tasks: ITask[]): void {
  const fmt = (d: string): string => formatDateOnly(d, 'MM/dd/yyyy', '');

  const rows = tasks.map(t => ({
    [strings.Export_ColTaskName]: t.title,
    [strings.Export_ColPhase]: t.phase,
    [strings.Export_ColStartDate]: fmt(t.startDate),
    [strings.Export_ColDueDate]: fmt(t.dueDate),
    [strings.Export_ColStatus]: t.status,
    [strings.Export_ColPriority]: t.priority,
    [strings.Export_ColAssignedTo]: t.assignedTo,
    [strings.Export_ColAssignedToEmail]: t.assignedToEmail,
    [strings.Export_ColPercentComplete]: t.percentComplete,
    [strings.Export_ColIsMilestone]: t.isMilestone ? strings.Export_Yes : strings.Export_No,
    [strings.Export_ColDescription]: t.description,
    [strings.Export_ColNotes]: t.notes,
  }));

  // json_to_sheet([]) produces a sheet with no header row at all — an
  // empty project would otherwise export a completely blank file.
  const ws = rows.length > 0
    ? XLSX.utils.json_to_sheet(rows)
    : XLSX.utils.aoa_to_sheet([TASK_EXPORT_HEADERS]);

  // Auto-width
  const headers = rows.length > 0 ? Object.keys(rows[0]) : TASK_EXPORT_HEADERS;
  ws['!cols'] = headers.map(h => ({
    wch: Math.max(h.length + 2, ...rows.map(r => String((r as Record<string, unknown>)[h] ?? '').length + 1)),
  }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, strings.Export_SheetNameTasks);
  XLSX.writeFile(wb, `${safeFileName(project.title)}${strings.Export_ExcelFileSuffix}`);
}

// ─── Gantt image export ───────────────────────────────────────────────────────

const LEFT_W = 300;
const TITLE_H = 48;
const HEADER_H = 56;
const BAR_H = 24;

function escXml(s: string): string {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Created/Modified are real timestamps — parse directly, not as date-only.
function parseTimestamp(s: string): Date | null {
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

interface IVisibleRow { type: 'task' | 'phase'; task?: ITask; phase?: string; }

function buildRows(tasks: ITask[]): IVisibleRow[] {
  const rows: IVisibleRow[] = [];
  const byPhase = new Map<string, ITask[]>();
  const noPhase: ITask[] = [];
  const subMap = new Map<number, ITask[]>();
  const ids = new Set(tasks.map(t => t.id));
  // Sub-tasks whose parent is missing are promoted to top level so they
  // never disappear from the export.
  const isSubTask = (t: ITask): boolean => !!t.parentTaskId && ids.has(t.parentTaskId);

  tasks.filter(isSubTask).forEach(t => {
    if (!subMap.has(t.parentTaskId!)) subMap.set(t.parentTaskId!, []);
    subMap.get(t.parentTaskId!)!.push(t);
  });

  tasks.filter(t => !isSubTask(t)).forEach(t => {
    if (t.phase) {
      if (!byPhase.has(t.phase)) byPhase.set(t.phase, []);
      byPhase.get(t.phase)!.push(t);
    } else {
      noPhase.push(t);
    }
  });

  const add = (task: ITask): void => {
    rows.push({ type: 'task', task });
    (subMap.get(task.id) || []).forEach(c => rows.push({ type: 'task', task: c }));
  };

  byPhase.forEach((pt, phase) => { rows.push({ type: 'phase', phase }); pt.forEach(add); });
  noPhase.forEach(add);
  return rows;
}

function taskDisplayColor(task: ITask, settings: IGanttDisplaySettings): string {
  const statusColor = STATUS_COLORS[task.status] || DEFAULT_COLOR;
  // A free-text task color that isn't a valid color falls back to the status color.
  if (task.color) return safeColor(task.color, statusColor);
  if (settings.colorBy === 'priority') return safeColor(PRIORITY_COLORS[task.priority], statusColor);
  if (settings.colorBy === 'phase' && task.phase) return safeColor(phaseColor(task.phase), statusColor);
  if (settings.colorBy === 'health') return safeColor(healthColor(computeTaskHealth(task)), statusColor);
  return safeColor(statusColor);
}

export function renderGanttSVG(
  project: IProject,
  tasks: ITask[],
  settings: IGanttDisplaySettings
): string {
  const ROW_H = settings.rowHeight;
  const theme = HEADER_THEME_COLORS[settings.headerTheme];
  const projectColorSafe = safeColor(project.color);

  // Date range
  const today = todayLocalMidnight();
  const allDates: Date[] = [addDays(today, -14)];
  tasks.forEach(t => {
    const s = parseDateOnly(t.startDate); const e = parseDateOnly(t.dueDate);
    if (s) allDates.push(addDays(s, -7));
    if (e) allDates.push(addDays(e, 14));
  });
  const rangeStart = startOfMonth(allDates.reduce((a, b) => a < b ? a : b));
  const rangeEnd = endOfMonth(allDates.reduce((a, b) => a > b ? a : b));
  const totalDays = differenceInCalendarDays(rangeEnd, rangeStart) + 1;

  // Day width based on total days (auto-fit)
  const DAY_W = Math.max(3, Math.min(40, Math.round(1200 / totalDays)));
  const timelineW = totalDays * DAY_W;
  const totalW = LEFT_W + timelineW;

  const visRows = buildRows(tasks);
  const bodyH = visRows.length * ROW_H + 20;
  const totalH = TITLE_H + HEADER_H + bodyH;

  const dateToX = (d: Date): number => differenceInCalendarDays(d, rangeStart) * DAY_W;
  const todayX = LEFT_W + dateToX(today) + DAY_W / 2;

  // Project-relative week calculation
  const projectStart = tasks.reduce<Date | null>((acc, t) => {
    const s = parseDateOnly(t.startDate);
    return s && (!acc || s < acc) ? s : acc;
  }, null) || today;
  const projectWeekStart = startOfWeek(projectStart, { weekStartsOn: 1 });

  const getWeekLabel = (weekStartDate: Date): string => {
    if (settings.weekLabel === 'dates') return format(weekStartDate, 'MMM d');
    if (settings.weekLabel === 'project') {
      const diff = differenceInCalendarDays(weekStartDate, projectWeekStart);
      const wn = Math.floor(diff / 7) + 1;
      return wn > 0 ? `W${wn}` : `W1`;
    }
    return `W${getISOWeek(weekStartDate)}`;
  };

  // Build month bands
  const months: { label: string; x: number; width: number }[] = [];
  let cur = startOfMonth(rangeStart);
  while (cur <= rangeEnd) {
    const ms = max([cur, rangeStart]);
    const me = min([endOfMonth(cur), rangeEnd]);
    months.push({
      label: format(cur, 'MMM yyyy'),
      x: dateToX(ms),
      width: (differenceInCalendarDays(me, ms) + 1) * DAY_W,
    });
    cur = addMonths(cur, 1);
  }

  // Build week bands
  const weeks: { label: string; x: number; width: number; isCurrent: boolean }[] = [];
  const thisWeekStart = startOfWeek(today, { weekStartsOn: 1 });
  let wcur = startOfWeek(rangeStart, { weekStartsOn: 1 });
  while (wcur <= rangeEnd) {
    const ws = max([wcur, rangeStart]);
    const we = min([addDays(wcur, 6), rangeEnd]);
    weeks.push({
      label: getWeekLabel(wcur),
      x: dateToX(ws),
      width: (differenceInCalendarDays(we, ws) + 1) * DAY_W,
      isCurrent: wcur.getTime() === thisWeekStart.getTime(),
    });
    wcur = addWeeks(wcur, 1);
  }

  // Weekend columns
  const weekendRects: string[] = [];
  if (settings.showWeekends && DAY_W >= 5) {
    let wd = rangeStart;
    while (wd <= rangeEnd) {
      if (isWeekend(wd)) {
        weekendRects.push(
          `<rect x="${LEFT_W + dateToX(wd)}" y="${TITLE_H + HEADER_H}" width="${DAY_W}" height="${bodyH}" fill="#F8F7F6"/>`
        );
      }
      wd = addDays(wd, 1);
    }
  }

  // Task bars
  const bars: string[] = [];
  const arrows: string[] = [];
  const taskIndexMap = new Map<number, number>();
  visRows.forEach((row, i) => { if (row.type === 'task' && row.task) taskIndexMap.set(row.task.id, i); });

  visRows.forEach((row, i) => {
    const y0 = TITLE_H + HEADER_H + i * ROW_H;
    if (row.type === 'phase') {
      bars.push(
        `<rect x="0" y="${y0}" width="${LEFT_W}" height="${ROW_H}" fill="#F3F2F1"/>`,
        `<rect x="${LEFT_W}" y="${y0}" width="${timelineW}" height="${ROW_H}" fill="#F8F7F6"/>`,
        `<text x="14" y="${y0 + ROW_H / 2 + 4}" font-family="Segoe UI,sans-serif" font-size="11" font-weight="700" fill="#605E5C" text-transform="uppercase" letter-spacing="0.5">${escXml((row.phase || '').toUpperCase())}</text>`
      );
      return;
    }

    const task = row.task!;
    const color = taskDisplayColor(task, settings);
    const sd = parseDateOnly(task.startDate); const ed = parseDateOnly(task.dueDate);
    const isChild = !!task.parentTaskId;
    const nameX = isChild ? 28 : 16;

    // Left panel row
    bars.push(
      `<rect x="0" y="${y0}" width="${LEFT_W}" height="${ROW_H}" fill="${i % 2 === 0 ? '#FFFFFF' : '#FAFAFA'}"/>`,
      `<line x1="0" y1="${y0 + ROW_H}" x2="${LEFT_W}" y2="${y0 + ROW_H}" stroke="#F3F2F1" stroke-width="1"/>`,
      `<circle cx="${isChild ? 20 : 10}" cy="${y0 + ROW_H / 2}" r="4" fill="${color}"/>`,
      `<text x="${nameX + 8}" y="${y0 + ROW_H / 2 + 4}" font-family="Segoe UI,sans-serif" font-size="12" fill="#323130">${escXml(task.title.substring(0, 38))}</text>`
    );

    if (!sd || !ed) return;

    const bx = LEFT_W + dateToX(sd);
    const bw = Math.max(4, (differenceInCalendarDays(ed, sd) + 1) * DAY_W);
    const by = y0 + (ROW_H - BAR_H) / 2;
    const pw = bw * (task.percentComplete / 100);
    const gradId = `g${task.id}`;

    if (task.isMilestone) {
      const mx = LEFT_W + dateToX(sd) + DAY_W / 2;
      const my = y0 + ROW_H / 2;
      const ms = 9;
      bars.push(
        `<polygon points="${mx},${my - ms} ${mx + ms},${my} ${mx},${my + ms} ${mx - ms},${my}" fill="${color}" stroke="white" stroke-width="1.5"/>`
      );
      return;
    }

    if (settings.barStyle === 'gradient') {
      bars.push(
        `<defs><linearGradient id="${gradId}" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="${color}" stop-opacity="1"/><stop offset="100%" stop-color="${color}" stop-opacity="0.65"/></linearGradient></defs>`,
        `<rect x="${bx}" y="${by}" width="${bw}" height="${BAR_H}" rx="4" fill="${color}22"/>`,
        pw > 0 ? `<rect x="${bx}" y="${by}" width="${pw}" height="${BAR_H}" rx="4" fill="url(#${gradId})"/>` : '',
      );
    } else {
      bars.push(
        `<rect x="${bx}" y="${by}" width="${bw}" height="${BAR_H}" rx="4" fill="${color}28"/>`,
        pw > 0 ? `<rect x="${bx}" y="${by}" width="${pw}" height="${BAR_H}" rx="4" fill="${color}"/>` : '',
      );
    }

    if (settings.showProgressText && task.percentComplete > 0 && bw > 40) {
      bars.push(
        `<text x="${bx + 6}" y="${by + BAR_H / 2 + 4}" font-family="Segoe UI,sans-serif" font-size="10" font-weight="700" fill="${pw > bw * 0.45 ? '#fff' : color}">${task.percentComplete}%</text>`
      );
    }

    if (settings.showAssignee && task.assignedTo) {
      bars.push(
        `<text x="${bx + bw + 6}" y="${by + BAR_H / 2 + 4}" font-family="Segoe UI,sans-serif" font-size="10" fill="#605E5C">${escXml(task.assignedTo)}</text>`
      );
    }

    // Right panel row line
    bars.push(`<line x1="${LEFT_W}" y1="${y0 + ROW_H}" x2="${LEFT_W + timelineW}" y2="${y0 + ROW_H}" stroke="#F3F2F1" stroke-width="1"/>`);

    // Dependency arrows
    if (settings.showDependencies) {
      task.dependencies.forEach(depId => {
        const depIdx = taskIndexMap.get(depId);
        if (depIdx === undefined) return;
        const depTask = visRows[depIdx]?.task;
        if (!depTask) return;
        const depEnd = parseDateOnly(depTask.dueDate);
        if (!depEnd) return;
        const fromX = LEFT_W + dateToX(depEnd) + DAY_W;
        const fromY = TITLE_H + HEADER_H + depIdx * ROW_H + ROW_H / 2;
        const toX = bx;
        const toY = y0 + ROW_H / 2;
        const midX = fromX + (toX - fromX) / 2;
        arrows.push(
          `<path d="M${fromX} ${fromY} C${midX} ${fromY},${midX} ${toY},${toX} ${toY}" fill="none" stroke="#8A8886" stroke-width="1.5" marker-end="url(#arr)"/>`
        );
      });
    }
  });

  // Divider line between left panel and timeline
  const dividerLine = `<line x1="${LEFT_W}" y1="${TITLE_H}" x2="${LEFT_W}" y2="${totalH}" stroke="#EDEBE9" stroke-width="2"/>`;

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${totalW}" height="${totalH}" style="background:white;font-family:'Segoe UI',Arial,sans-serif">
  <defs>
    <marker id="arr" markerWidth="8" markerHeight="6" refX="7" refY="3" orient="auto">
      <polygon points="0 0,8 3,0 6" fill="#8A8886"/>
    </marker>
  </defs>

  <!-- ── Title bar ─────────────────────────────────────────────────── -->
  <rect width="${totalW}" height="${TITLE_H}" fill="${projectColorSafe}"/>
  <circle cx="24" cy="${TITLE_H / 2}" r="7" fill="white" opacity="0.25"/>
  <text x="38" y="${TITLE_H / 2 + 6}" font-size="17" font-weight="700" fill="white">${escXml(project.title)}</text>
  <text x="${totalW - 12}" y="${TITLE_H / 2 + 5}" font-size="11" fill="rgba(255,255,255,0.7)" text-anchor="end">${escXml(formatString(strings.Export_ExportedDateLabel, { date: format(today, 'MMM d, yyyy') }))}</text>

  <!-- ── Left panel header ─────────────────────────────────────────── -->
  <rect y="${TITLE_H}" width="${LEFT_W}" height="${HEADER_H}" fill="${theme.bg}"/>
  <text x="16" y="${TITLE_H + HEADER_H / 2 + 4}" font-size="11" font-weight="600" fill="${theme.subtext}" letter-spacing="0.5">${escXml(strings.Export_TaskNameColumnLabel)}</text>

  <!-- ── Timeline header ───────────────────────────────────────────── -->
  <rect x="${LEFT_W}" y="${TITLE_H}" width="${timelineW}" height="${HEADER_H}" fill="${theme.bg}"/>

  <!-- Month bands (top half of header) -->
  ${months.map(m => `
    <line x1="${LEFT_W + m.x}" y1="${TITLE_H}" x2="${LEFT_W + m.x}" y2="${TITLE_H + 28}" stroke="rgba(255,255,255,0.15)" stroke-width="1"/>
    <text x="${LEFT_W + m.x + 6}" y="${TITLE_H + 18}" font-size="12" font-weight="600" fill="${theme.text}">${escXml(m.label)}</text>
  `).join('')}

  <!-- Week bands (bottom half of header) -->
  ${weeks.map(w => `
    ${w.isCurrent ? `<rect x="${LEFT_W + w.x}" y="${TITLE_H + 28}" width="${w.width}" height="28" fill="rgba(255,215,0,0.15)"/>` : ''}
    <line x1="${LEFT_W + w.x}" y1="${TITLE_H + 28}" x2="${LEFT_W + w.x}" y2="${TITLE_H + HEADER_H}" stroke="rgba(255,255,255,0.1)" stroke-width="1"/>
    <text x="${LEFT_W + w.x + 4}" y="${TITLE_H + HEADER_H - 9}" font-size="10" fill="${w.isCurrent ? '#FFD700' : theme.subtext}" font-weight="${w.isCurrent ? '700' : '400'}">${escXml(w.label)}</text>
  `).join('')}

  <!-- ── Body ──────────────────────────────────────────────────────── -->
  <rect y="${TITLE_H + HEADER_H}" width="${totalW}" height="${bodyH}" fill="white"/>

  ${weekendRects.join('\n  ')}
  ${bars.join('\n  ')}
  ${arrows.join('\n  ')}
  ${dividerLine}

  <!-- Today line -->
  <line x1="${todayX}" y1="${TITLE_H + HEADER_H}" x2="${todayX}" y2="${totalH}" stroke="#D13438" stroke-width="2" stroke-dasharray="4,3"/>
  <circle cx="${todayX}" cy="${TITLE_H + HEADER_H}" r="4" fill="#D13438"/>
</svg>`;
}

export function downloadPNG(svgString: string, filename: string, scale: number = 2): Promise<void> {
  filename = safeFileName(filename);
  return svgToCanvas(svgString, scale).then(canvas => new Promise<void>((resolve, reject) => {
    canvas.toBlob(pngBlob => {
      if (!pngBlob) {
        reject(new Error(strings.Export_ImageGenerationFailed));
        return;
      }
      const pngUrl = URL.createObjectURL(pngBlob);
      const a = document.createElement('a');
      a.href = pngUrl;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(pngUrl);
      resolve();
    }, 'image/png');
  }));
}

// ─── PowerPoint export ────────────────────────────────────────────────────────

// Browsers cap canvas dimensions around 16,384px (varies by browser/GPU); a
// multi-year project at the default scale can exceed that (e.g. ~2,900 days
// x 3px x 2 scale ~= 17,400px), silently producing a null/blank image.
const MAX_CANVAS_DIMENSION = 16000;

// Shared SVG→canvas helper used by both downloadPNG and svgToPngDataUrl.
function svgToCanvas(svgString: string, scale: number): Promise<HTMLCanvasElement> {
  return new Promise((resolve, reject) => {
    const blob = new Blob([svgString], { type: 'image/svg+xml' });
    const svgUrl = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = (): void => {
      const effectiveScale = Math.min(scale, MAX_CANVAS_DIMENSION / img.width, MAX_CANVAS_DIMENSION / img.height);
      const canvas = document.createElement('canvas');
      canvas.width = img.width * effectiveScale;
      canvas.height = img.height * effectiveScale;
      const ctx = canvas.getContext('2d');
      if (!ctx) {
        URL.revokeObjectURL(svgUrl);
        reject(new Error(strings.Export_DrawingContextFailed));
        return;
      }
      ctx.scale(effectiveScale, effectiveScale);
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, img.width, img.height);
      ctx.drawImage(img, 0, 0);
      URL.revokeObjectURL(svgUrl);
      resolve(canvas);
    };
    img.onerror = (): void => {
      URL.revokeObjectURL(svgUrl);
      reject(new Error(strings.Export_SvgRenderFailed));
    };
    img.src = svgUrl;
  });
}

function svgToPngDataUrl(svgString: string, scale: number = 2): Promise<string> {
  return svgToCanvas(svgString, scale).then(canvas => canvas.toDataURL('image/png'));
}

// PowerPoint wants 'RRGGBB' with no '#'; validated so a stray value can't corrupt the XML.
function hex(color: string): string {
  return safeColor(color).slice(1).toUpperCase();
}

export async function exportToPowerPoint(
  project: IProject,
  tasks: ITask[],
  settings: IGanttDisplaySettings
): Promise<void> {
  const today = new Date();
  const fmt = (d: string): string => formatDateOnly(d, 'MMM d, yyyy');
  const projectColor = hex(project.color);

  const ganttDataUrl = await svgToPngDataUrl(renderGanttSVG(project, tasks, settings), 2);

  // Compute summary stats
  const totalCount = tasks.length;
  const byStatus: Record<string, number> = {
    'Not Started': 0, 'In Progress': 0, 'Completed': 0, 'On Hold': 0, 'Cancelled': 0,
  };
  tasks.forEach(t => { if (byStatus[t.status] !== undefined) byStatus[t.status]++; });
  const overallPct = totalCount > 0
    ? Math.round(tasks.reduce((s, t) => s + t.percentComplete, 0) / totalCount)
    : 0;

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE'; // 13.33" × 7.5"

  // ── Slide 1: Cover ─────────────────────────────────────────────────────────
  const cover = pptx.addSlide();
  cover.background = { color: projectColor };

  // White lower panel
  cover.addShape(pptx.ShapeType.rect, {
    x: 0, y: 4.6, w: 13.33, h: 2.9,
    fill: { color: 'FFFFFF' },
    line: { color: 'FFFFFF', width: 0 },
  });

  cover.addText(strings.Export_ProjectReportLabel, {
    x: 0.65, y: 0.9, w: 12, h: 0.45,
    fontSize: 11,
    color: 'FFFFFF',
    charSpacing: 3,
    transparency: 30,
  });

  cover.addText(project.title, {
    x: 0.65, y: 1.35, w: 12, h: 1.6,
    fontSize: 42,
    color: 'FFFFFF',
    bold: true,
  });

  cover.addText(formatString(strings.Export_StatusLine, { status: project.status }), {
    x: 0.65, y: 3.1, w: 6, h: 0.45,
    fontSize: 16,
    color: 'FFFFFF',
    transparency: 15,
  });

  const dateRange = (project.startDate || project.dueDate)
    ? `${fmt(project.startDate)}${strings.Export_DateRangeSeparator}${fmt(project.dueDate)}`
    : '';
  if (dateRange) {
    cover.addText(dateRange, {
      x: 0.65, y: 3.6, w: 10, h: 0.4,
      fontSize: 14,
      color: 'FFFFFF',
      transparency: 30,
    });
  }

  if (project.description) {
    cover.addText(project.description, {
      x: 0.65, y: 4.75, w: 12, h: 1.0,
      fontSize: 13,
      color: '323130',
    });
  }

  if (project.projectManager) {
    cover.addText(formatString(strings.Export_ProjectManagerLine, { name: project.projectManager }), {
      x: 0.65, y: 5.8, w: 8, h: 0.35,
      fontSize: 12,
      color: '605E5C',
    });
  }

  cover.addText(formatString(strings.Export_GeneratedDateLine, { date: format(today, 'MMMM d, yyyy') }), {
    x: 0, y: 7.15, w: 13.15, h: 0.3,
    fontSize: 10,
    color: '605E5C',
    align: 'right',
  });

  // ── Slide 2: Project Summary ───────────────────────────────────────────────
  const summary = pptx.addSlide();
  summary.background = { color: 'FFFFFF' };

  // Header bar
  summary.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 13.33, h: 0.85,
    fill: { color: projectColor },
    line: { color: projectColor, width: 0 },
  });
  summary.addText(strings.Export_ProjectSummaryTitle, {
    x: 0.4, y: 0, w: 9, h: 0.85,
    fontSize: 22, color: 'FFFFFF', bold: true, valign: 'middle',
  });
  summary.addText(project.title, {
    x: 0, y: 0, w: 13.0, h: 0.85,
    fontSize: 13, color: 'FFFFFF', align: 'right', valign: 'middle', transparency: 35,
  });

  // Stat boxes
  const statItems = [
    { label: strings.Export_StatTotalTasks,  value: totalCount,                   bg: 'F3F2F1', fg: '323130' },
    { label: strings.Export_StatCompleted,   value: byStatus['Completed'],        bg: 'F1FAF1', fg: '107C10' },
    { label: strings.Export_StatInProgress,  value: byStatus['In Progress'],      bg: 'EFF6FC', fg: '0078D4' },
    { label: strings.Export_StatOnHold,      value: byStatus['On Hold'],          bg: 'FFF4EC', fg: 'CA5010' },
    { label: strings.Export_StatNotStarted,  value: byStatus['Not Started'],      bg: 'F3F2F1', fg: '605E5C' },
  ];

  const BOX_W = 2.3; const BOX_H = 1.4; const BOX_Y = 1.1; const GAP = 0.165;
  const BOX_START = (13.33 - (statItems.length * BOX_W + (statItems.length - 1) * GAP)) / 2;

  statItems.forEach((item, i) => {
    const x = BOX_START + i * (BOX_W + GAP);
    summary.addShape(pptx.ShapeType.rect, {
      x, y: BOX_Y, w: BOX_W, h: BOX_H,
      fill: { color: item.bg },
      line: { color: 'EDEBE9', width: 1 },
      rectRadius: 0.05,
    });
    summary.addText(String(item.value), {
      x, y: BOX_Y + 0.08, w: BOX_W, h: 0.82,
      fontSize: 34, color: item.fg, bold: true, align: 'center', valign: 'middle',
    });
    summary.addText(item.label, {
      x, y: BOX_Y + 0.95, w: BOX_W, h: 0.38,
      fontSize: 11, color: '605E5C', align: 'center',
    });
  });

  // Overall progress bar
  const BAR_Y = 2.85;
  const overallProgressPercentText = formatString(strings.Export_OverallProgressPercent, { percent: overallPct });
  summary.addText(`${strings.Export_OverallProgressLine}: ${overallProgressPercentText}`, {
    x: 0.5, y: BAR_Y, w: 8, h: 0.35,
    fontSize: 13, bold: true, color: '323130',
  });
  summary.addShape(pptx.ShapeType.rect, {
    x: 0.5, y: BAR_Y + 0.42, w: 12.33, h: 0.38,
    fill: { color: 'EDEBE9' }, line: { color: 'EDEBE9', width: 0 },
  });
  if (overallPct > 0) {
    summary.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: BAR_Y + 0.42, w: 12.33 * overallPct / 100, h: 0.38,
      fill: { color: projectColor }, line: { color: projectColor, width: 0 },
    });
  }
  summary.addText(overallProgressPercentText, {
    x: 12.85, y: BAR_Y, w: 0.8, h: 0.35,
    fontSize: 13, bold: true, color: projectColor, align: 'right',
  });

  // Status breakdown (two columns)
  const TBL_Y = 3.75;
  summary.addText(strings.Export_StatusBreakdownTitle, {
    x: 0.5, y: TBL_Y, w: 6, h: 0.35,
    fontSize: 12, bold: true, color: '323130',
  });

  const statusRows = (
    ['Completed', 'In Progress', 'Not Started', 'On Hold', 'Cancelled'] as const
  ).map(label => ({ label, dotColor: hex(STATUS_COLORS[label]) }));

  statusRows.forEach((s, i) => {
    const col = Math.floor(i / 3);
    const row = i % 3;
    const cx = 0.5 + col * 6.2;
    const cy = TBL_Y + 0.45 + row * 0.45;
    const cnt = byStatus[s.label] ?? 0;
    const pct = totalCount > 0 ? Math.round(cnt / totalCount * 100) : 0;

    summary.addShape(pptx.ShapeType.ellipse, {
      x: cx, y: cy + 0.06, w: 0.2, h: 0.2,
      fill: { color: s.dotColor }, line: { color: s.dotColor, width: 0 },
    });
    summary.addText(formatString(strings.Export_StatusBreakdownRow, { status: s.label, count: cnt, percent: pct }), {
      x: cx + 0.28, y: cy, w: 5.6, h: 0.36,
      fontSize: 12, color: '323130',
    });
  });

  // ── Slide 3: Gantt Timeline ────────────────────────────────────────────────
  const gantt = pptx.addSlide();
  gantt.background = { color: 'FFFFFF' };

  gantt.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 13.33, h: 0.85,
    fill: { color: projectColor },
    line: { color: projectColor, width: 0 },
  });
  gantt.addText(strings.Export_GanttTimelineTitle, {
    x: 0.4, y: 0, w: 9, h: 0.85,
    fontSize: 22, color: 'FFFFFF', bold: true, valign: 'middle',
  });
  gantt.addText(project.title, {
    x: 0, y: 0, w: 13.0, h: 0.85,
    fontSize: 13, color: 'FFFFFF', align: 'right', valign: 'middle', transparency: 35,
  });

  if (ganttDataUrl) {
    gantt.addImage({
      data: ganttDataUrl,
      x: 0.1, y: 0.95, w: 13.13, h: 6.45,
      sizing: { type: 'contain', w: 13.13, h: 6.45 },
    });
  }

  // ── Slide 4: Summary & Recent Activity ────────────────────────────────────
  const activity = pptx.addSlide();
  activity.background = { color: 'FFFFFF' };

  // Header bar
  activity.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 13.33, h: 0.85,
    fill: { color: projectColor }, line: { color: projectColor, width: 0 },
  });
  activity.addText(strings.Export_SummaryRecentActivityTitle, {
    x: 0.4, y: 0, w: 9, h: 0.85,
    fontSize: 22, color: 'FFFFFF', bold: true, valign: 'middle',
  });
  activity.addText(project.title, {
    x: 0, y: 0, w: 13.0, h: 0.85,
    fontSize: 13, color: 'FFFFFF', align: 'right', valign: 'middle', transparency: 35,
  });

  // Vertical divider
  activity.addShape(pptx.ShapeType.rect, {
    x: 6.58, y: 0.95, w: 0.02, h: 6.4,
    fill: { color: 'EDEBE9' }, line: { color: 'EDEBE9', width: 0 },
  });

  // ── LEFT column: Project Overview ──────────────────────────────────────────
  const LX = 0.5; const LW = 5.85;

  activity.addText(strings.Export_ProjectOverviewLabel, {
    x: LX, y: 1.05, w: LW, h: 0.3,
    fontSize: 10, bold: true, color: '605E5C', charSpacing: 1.5,
  });

  if (project.description) {
    activity.addText(project.description, {
      x: LX, y: 1.45, w: LW, h: 2.0,
      fontSize: 13, color: '323130',
    });
  }

  const detailStartY = project.description ? 3.65 : 1.45;
  const details: { label: string; value: string }[] = [
    { label: strings.Export_DetailStatusLabel, value: project.status },
    ...(project.startDate ? [{ label: strings.Export_DetailStartDateLabel, value: fmt(project.startDate) }] : []),
    ...(project.dueDate   ? [{ label: strings.Export_DetailDueDateLabel,   value: fmt(project.dueDate)   }] : []),
    ...(project.projectManager ? [{ label: strings.Export_DetailProjectManagerLabel, value: project.projectManager }] : []),
  ];

  details.forEach((d, i) => {
    const dy = detailStartY + i * 0.47;
    activity.addText(`${d.label}:`, {
      x: LX, y: dy, w: 1.9, h: 0.36,
      fontSize: 12, bold: true, color: '323130',
    });
    activity.addText(d.value, {
      x: LX + 1.95, y: dy, w: LW - 1.95, h: 0.36,
      fontSize: 12, color: '605E5C',
    });
  });

  // ── RIGHT column: Past 7 Days ──────────────────────────────────────────────
  const RX = 6.75; const RW = 6.3;

  activity.addText(strings.Export_Past7DaysLabel, {
    x: RX, y: 1.05, w: RW, h: 0.3,
    fontSize: 10, bold: true, color: '605E5C', charSpacing: 1.5,
  });

  const weekAgo = addDays(today, -7);
  const recentTasks = tasks.filter(t => {
    const mod = parseTimestamp(t.modified);
    return mod !== null && mod >= weekAgo;
  });

  const completedThisWeek = recentTasks.filter(t => t.status === 'Completed');
  const updatedThisWeek   = recentTasks.filter(t => t.status !== 'Completed');

  let ry = 1.45;
  const MAX_PER_SECTION = 4;

  if (recentTasks.length === 0) {
    activity.addText(strings.Export_NoActivityMessage, {
      x: RX, y: ry, w: RW, h: 0.4,
      fontSize: 12, color: '8A8886', italic: true,
    });
  } else {
    if (completedThisWeek.length > 0) {
      // Section header chip
      activity.addShape(pptx.ShapeType.rect, {
        x: RX, y: ry, w: RW, h: 0.33,
        fill: { color: 'F1FAF1' }, line: { color: 'C8E6C9', width: 1 }, rectRadius: 0.03,
      });
      activity.addText(formatString(strings.Export_CompletedThisWeekHeader, { count: completedThisWeek.length }), {
        x: RX + 0.15, y: ry, w: RW - 0.3, h: 0.33,
        fontSize: 11, bold: true, color: '107C10', valign: 'middle',
      });
      ry += 0.4;

      completedThisWeek.slice(0, MAX_PER_SECTION).forEach(t => {
        const created = parseTimestamp(t.created);
        const isNew = created !== null && created >= weekAgo;
        const titleRuns = isNew
          ? [{ text: strings.Export_NewBadge, options: { color: '107C10', bold: true, fontSize: 9 } },
             { text: t.title, options: { color: '323130', fontSize: 12 } }]
          : [{ text: t.title, options: { color: '323130', fontSize: 12 } }];

        activity.addShape(pptx.ShapeType.ellipse, {
          x: RX + 0.1, y: ry + 0.1, w: 0.15, h: 0.15,
          fill: { color: '107C10' }, line: { color: '107C10', width: 0 },
        });
        activity.addText(titleRuns, {
          x: RX + 0.34, y: ry, w: RW - 0.34, h: 0.26, valign: 'middle',
        });
        activity.addText(formatString(strings.Export_CompletedPercentLine, { percent: t.percentComplete }), {
          x: RX + 0.34, y: ry + 0.27, w: RW - 0.34, h: 0.18,
          fontSize: 9, color: '8A8886',
        });
        ry += 0.46;
      });

      if (completedThisWeek.length > MAX_PER_SECTION) {
        activity.addText(formatString(strings.Export_MoreItemsSuffix, { count: completedThisWeek.length - MAX_PER_SECTION }), {
          x: RX + 0.34, y: ry, w: RW, h: 0.3,
          fontSize: 11, color: '8A8886', italic: true,
        });
        ry += 0.32;
      }
      ry += 0.15;
    }

    if (updatedThisWeek.length > 0) {
      activity.addShape(pptx.ShapeType.rect, {
        x: RX, y: ry, w: RW, h: 0.33,
        fill: { color: 'EFF6FC' }, line: { color: 'BFDBF7', width: 1 }, rectRadius: 0.03,
      });
      activity.addText(`${strings.Export_InProgressUpdatedHeader}  (${updatedThisWeek.length})`, {
        x: RX + 0.15, y: ry, w: RW - 0.3, h: 0.33,
        fontSize: 11, bold: true, color: '0078D4', valign: 'middle',
      });
      ry += 0.4;

      updatedThisWeek.slice(0, MAX_PER_SECTION).forEach(t => {
        const created = parseTimestamp(t.created);
        const isNew = created !== null && created >= weekAgo;
        const dotColor = hex(STATUS_COLORS[t.status] || STATUS_COLORS['Not Started']);
        const titleRuns = isNew
          ? [{ text: 'NEW  ', options: { color: dotColor, bold: true, fontSize: 9 } },
             { text: t.title, options: { color: '323130', fontSize: 12 } }]
          : [{ text: t.title, options: { color: '323130', fontSize: 12 } }];

        activity.addShape(pptx.ShapeType.ellipse, {
          x: RX + 0.1, y: ry + 0.1, w: 0.15, h: 0.15,
          fill: { color: dotColor }, line: { color: dotColor, width: 0 },
        });
        activity.addText(titleRuns, {
          x: RX + 0.34, y: ry, w: RW - 0.34, h: 0.26, valign: 'middle',
        });
        activity.addText(formatString(strings.Export_StatusPercentLine, { status: t.status, percent: t.percentComplete }), {
          x: RX + 0.34, y: ry + 0.27, w: RW - 0.34, h: 0.18,
          fontSize: 9, color: '8A8886',
        });
        ry += 0.46;
      });

      if (updatedThisWeek.length > MAX_PER_SECTION) {
        activity.addText(formatString(strings.Export_MoreItemsSuffix, { count: updatedThisWeek.length - MAX_PER_SECTION }), {
          x: RX + 0.34, y: ry, w: RW, h: 0.3,
          fontSize: 11, color: '8A8886', italic: true,
        });
      }
    }
  }

  await pptx.writeFile({ fileName: formatString(strings.Export_ProjectReportFileName, { projectName: safeFileName(project.title) }) });
}

// ─── Portfolio exports ────────────────────────────────────────────────────────

function portfolioHealthLabel(h: string): string {
  return healthLabel(h as Parameters<typeof healthLabel>[0]);
}

function portfolioHealthHex(h: string): string {
  const map: Record<string, string> = { complete: '107C10', 'on-track': '0078D4', 'at-risk': 'CA5010', overdue: 'D13438' };
  return map[h] ?? '323130';
}

const PORTFOLIO_HEADERS = [
  strings.Export_TableProject, strings.Export_TableStatus, strings.Export_TableHealth,
  strings.Export_StatTotalTasks, strings.Export_StatCompleted, strings.Export_StatInProgress,
  strings.Export_TableAtRisk, strings.Export_TableOverdue, strings.Export_TablePctDone,
  strings.Export_ColStart, strings.Export_ColDue, strings.Export_ColDescription,
];

function portfolioRows(
  projects: IProject[],
  statsMap: Map<number, IProjectTaskStats> | null,
  fmt: (d: string) => string
): Array<Record<string, string | number>> {
  return projects.map(p => {
    const s = statsMap?.get(p.id);
    const ok = s && !s.statsError;
    return {
      [strings.Export_TableProject]:     p.title,
      [strings.Export_TableStatus]:      p.status,
      [strings.Export_TableHealth]:      s ? (s.statsError ? strings.Export_UnavailableLabel : portfolioHealthLabel(s.health)) : '—',
      [strings.Export_StatTotalTasks]:   ok ? s!.totalTasks : '—',
      [strings.Export_StatCompleted]:    ok ? s!.completedCount : '—',
      [strings.Export_StatInProgress]:   ok ? s!.inProgressCount : '—',
      [strings.Export_TableAtRisk]:      ok ? s!.atRiskCount : '—',
      [strings.Export_TableOverdue]:     ok ? s!.overdueCount : '—',
      [strings.Export_TablePctDone]:     ok ? `${s!.overallPct}%` : '—',
      [strings.Export_ColStart]:         fmt(p.startDate || (ok ? s!.earliestStart : '') || ''),
      [strings.Export_ColDue]:           fmt(p.dueDate   || (ok ? s!.latestDue    : '') || ''),
      [strings.Export_ColDescription]:   p.description,
    };
  });
}

export function exportPortfolioToExcel(
  projects: IProject[],
  statsMap: Map<number, IProjectTaskStats> | null
): void {
  const fmt = (d: string): string => formatDateOnly(d, 'MM/dd/yyyy', '');
  const rows = portfolioRows(projects, statsMap, fmt);

  // json_to_sheet([]) has no header row at all, so with no projects the file
  // would be completely blank — emit the header row explicitly instead.
  const ws = rows.length > 0
    ? XLSX.utils.json_to_sheet(rows)
    : XLSX.utils.aoa_to_sheet([PORTFOLIO_HEADERS]);
  const headers = rows.length > 0 ? Object.keys(rows[0]) : PORTFOLIO_HEADERS;
  ws['!cols'] = headers.map(h => ({
    wch: Math.max(h.length + 2, ...rows.map(r => String(r[h] ?? '').length + 1)),
  }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, strings.Export_SheetNamePortfolio);
  XLSX.writeFile(wb, strings.Export_PortfolioSummaryFileName);
}

export async function exportPortfolioToPowerPoint(
  projects: IProject[],
  statsMap: Map<number, IProjectTaskStats> | null
): Promise<void> {
  const today = new Date();
  const fmt = (d: string): string => formatDateOnly(d, 'MMM d, yyyy');
  const ACCENT = '0078D4';

  // Aggregate health counts
  const healthCounts = { 'on-track': 0, 'at-risk': 0, overdue: 0, complete: 0 };
  if (statsMap) {
    statsMap.forEach(s => { if (!s.statsError && s.health in healthCounts) healthCounts[s.health as keyof typeof healthCounts]++; });
  }

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';

  // ── Slide 1: Cover ─────────────────────────────────────────────────────────
  const cover = pptx.addSlide();
  cover.background = { color: ACCENT };

  cover.addShape(pptx.ShapeType.rect, {
    x: 0, y: 4.6, w: 13.33, h: 2.9,
    fill: { color: 'FFFFFF' }, line: { color: 'FFFFFF', width: 0 },
  });

  cover.addText(strings.Export_PortfolioReportLabel, {
    x: 0.65, y: 0.9, w: 12, h: 0.45,
    fontSize: 11, color: 'FFFFFF', charSpacing: 3, transparency: 30,
  });
  cover.addText(strings.Export_PortfolioOverviewTitle, {
    x: 0.65, y: 1.35, w: 12, h: 1.4,
    fontSize: 42, color: 'FFFFFF', bold: true,
  });
  cover.addText(formatString(strings.Export_ProjectsCountLine, { count: projects.length }), {
    x: 0.65, y: 3.0, w: 6, h: 0.45,
    fontSize: 16, color: 'FFFFFF', transparency: 15,
  });

  // Health summary chips
  const summaryParts: string[] = [];
  if (healthCounts['on-track'])  summaryParts.push(formatString(strings.Export_HealthSummaryOnTrack, { count: healthCounts['on-track'] }));
  if (healthCounts['at-risk'])   summaryParts.push(formatString(strings.Export_HealthSummaryAtRisk, { count: healthCounts['at-risk'] }));
  if (healthCounts.overdue)      summaryParts.push(formatString(strings.Export_HealthSummaryOverdue, { count: healthCounts.overdue }));
  if (summaryParts.length) {
    cover.addText(summaryParts.join('  ·  '), {
      x: 0.65, y: 3.5, w: 12, h: 0.4,
      fontSize: 14, color: 'FFFFFF', transparency: 30,
    });
  }
  cover.addText(formatString(strings.Export_GeneratedDateLine, { date: format(today, 'MMMM d, yyyy') }), {
    x: 0, y: 7.15, w: 13.15, h: 0.3,
    fontSize: 10, color: '605E5C', align: 'right',
  });

  // ── Slide 2+: Project Summary Table (paginated) ────────────────────────────
  const COL_W = [3.0, 0.9, 1.0, 0.75, 0.75, 0.75, 0.75, 0.75, 0.75, 1.35];
  const HDR   = [
    strings.Export_TableProject, strings.Export_TableStatus, strings.Export_TableHealth, strings.Export_TableTotal,
    strings.Export_TableDone, strings.Export_TableActive, strings.Export_TableAtRisk, strings.Export_TableOverdue,
    strings.Export_TablePctDone, strings.Export_TableDueDate,
  ];

  const hdrRow = HDR.map((h, _i) => ({
    text: h,
    options: {
      bold: true, fontSize: 10, color: 'FFFFFF', align: 'center' as const,
      fill: { color: ACCENT },
      border: [{ type: 'solid' as const, pt: 1, color: '0065B3' }],
      colspan: 1,
    },
  }));

  const dataRows = projects.map(p => {
    const s = statsMap?.get(p.id);
    const ok = s && !s.statsError;
    const hLabel = s ? (s.statsError ? strings.Export_UnavailableLabel : portfolioHealthLabel(s.health)) : '—';
    const hHex   = ok ? portfolioHealthHex(s!.health) : '323130';
    const rowBg  = 'FFFFFF';

    const cell = (txt: string, opts: object = {}): object => ({
      text: txt,
      options: { fontSize: 10, color: '323130', align: 'center' as const, fill: { color: rowBg }, border: [{ type: 'solid' as const, pt: 1, color: 'EDEBE9' }], ...opts },
    });

    return [
      cell(p.title, { align: 'left' as const, bold: true }),
      cell(p.status),
      cell(hLabel, { color: hHex, bold: true }),
      cell(ok ? String(s!.totalTasks)    : '—'),
      cell(ok ? String(s!.completedCount): '—', { color: '107C10' }),
      cell(ok ? String(s!.inProgressCount):'—', { color: '0078D4' }),
      cell(ok ? String(s!.atRiskCount)   : '—', { color: 'CA5010' }),
      cell(ok ? String(s!.overdueCount)  : '—', { color: 'D13438' }),
      cell(ok ? `${s!.overallPct}%`      : '—'),
      cell(fmt(p.dueDate || (ok ? s!.latestDue : '') || '')),
    ];
  });

  // A slide fits about 12 project rows plus the header row; longer lists
  // continue on following slides (each repeats the header row) instead of
  // shrinking rows until they overflow the slide.
  const ROWS_PER_SLIDE = 12;
  const pageCount = Math.max(1, Math.ceil(dataRows.length / ROWS_PER_SLIDE));
  for (let page = 0; page < pageCount; page++) {
    const table = pptx.addSlide();
    table.background = { color: 'FFFFFF' };

    table.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: 13.33, h: 0.75,
      fill: { color: ACCENT }, line: { color: ACCENT, width: 0 },
    });
    table.addText(pageCount > 1 ? `${strings.Export_ProjectsSummaryHeader} (${page + 1}/${pageCount})` : strings.Export_ProjectsSummaryHeader, {
      x: 0.4, y: 0, w: 9, h: 0.75,
      fontSize: 20, color: 'FFFFFF', bold: true, valign: 'middle',
    });
    table.addText(`${formatString(strings.Export_ProjectsCountLine, { count: projects.length })}  ·  ${format(today, 'MMM d, yyyy')}`, {
      x: 0, y: 0, w: 13.0, h: 0.75,
      fontSize: 12, color: 'FFFFFF', align: 'right', valign: 'middle', transparency: 35,
    });

    const pageRows = dataRows.slice(page * ROWS_PER_SLIDE, (page + 1) * ROWS_PER_SLIDE);
    table.addTable([hdrRow, ...pageRows] as Parameters<typeof table.addTable>[0], {
      x: 0.15, y: 0.9,
      w: 13.0,
      colW: COL_W,
      rowH: 0.42,
      fontSize: 10,
    });
  }

  await pptx.writeFile({ fileName: strings.Export_PortfolioReportFileName });
}

// ─── CSV / iCalendar exports ──────────────────────────────────────────────────

// RFC 4180 quoting, plus a leading apostrophe on text that starts with a
// formula trigger so Excel doesn't execute user-entered "=..." cells.
function csvCell(value: string | number | boolean | null | undefined): string {
  if (value === null || value === undefined) return '';
  let text = String(value);
  if (typeof value === 'string' && /^[=+\-@\t\r]/.test(text)) text = `'${text}`;
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function downloadCsv(headers: string[], rows: Array<Array<string | number | boolean | null | undefined>>, fileName: string): void {
  const lines = [headers, ...rows].map(r => r.map(csvCell).join(','));
  // UTF-8 BOM so Excel opens accented/non-Latin text correctly.
  const blob = new Blob(['\uFEFF' + lines.join('\r\n') + '\r\n'], { type: 'text/csv;charset=utf-8' });
  downloadBlob(blob, safeFileName(fileName));
}

/** Download the project's tasks as a CSV file (dates as yyyy-MM-dd; dependencies as "12,15SS+2"). */
export function exportTasksCsv(project: IProject, tasks: ITask[]): void {
  const headers = [
    strings.Svc_ColId, ...TASK_EXPORT_HEADERS, strings.Svc_ColDependencies,
    strings.Svc_ColBaselineStart, strings.Svc_ColBaselineDue,
  ];
  const rows = tasks.map(t => [
    t.id, t.title, t.phase, t.startDate, t.dueDate, t.status, t.priority, t.assignedTo, t.assignedToEmail,
    t.percentComplete, t.isMilestone ? strings.Export_Yes : strings.Export_No, t.description, t.notes,
    serializeDependencies(t.dependencies || [], t.dependencyLinks),
    t.baselineStart || '', t.baselineDue || '',
  ]);
  downloadCsv(headers, rows, `${project.title}${strings.Svc_CsvFileSuffix}`);
}

/** Download the portfolio summary (one row per project) as a CSV file. */
export function exportProjectsCsv(
  projects: IProject[],
  statsMap: Map<number, IProjectTaskStats> | null
): void {
  const rows = portfolioRows(projects, statsMap, d => d);
  downloadCsv(PORTFOLIO_HEADERS, rows.map(r => PORTFOLIO_HEADERS.map(h => r[h])), strings.Svc_PortfolioCsvFileName);
}

// iCalendar text values escape backslash, semicolon, comma and newlines.
function icsText(v: string): string {
  return String(v || '').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n');
}

// Content lines are limited to 75 octets; fold longer ones with CRLF + space.
function icsFold(line: string): string {
  const parts: string[] = [];
  let rest = line;
  while (rest.length > 74) {
    parts.push(rest.slice(0, 74));
    rest = ` ${rest.slice(74)}`;
  }
  parts.push(rest);
  return parts.join('\r\n');
}

function icsDate(iso: string): string {
  return iso.replace(/-/g, '');
}

/** Download the project's milestones as an .ics calendar of all-day events. */
export function exportMilestonesIcs(project: IProject, tasks: ITask[]): void {
  const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d+/, '');
  const lines: string[] = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Smart Gantt//Milestones//EN',
    'CALSCALE:GREGORIAN',
    'METHOD:PUBLISH',
    `X-WR-CALNAME:${icsText(project.title)}`,
  ];
  tasks.filter(t => t.isMilestone).forEach(t => {
    const dayStr = t.dueDate || t.startDate;
    const day = parseDateOnly(dayStr);
    if (!day) return;
    lines.push(
      'BEGIN:VEVENT',
      `UID:${project.id}-${t.id}@smartgantt`,
      `DTSTAMP:${stamp}`,
      `DTSTART;VALUE=DATE:${icsDate(dateOnlyString(day))}`,
      `DTEND;VALUE=DATE:${icsDate(dateOnlyString(addDays(day, 1)))}`,
      `SUMMARY:${icsText(t.title)}`,
    );
    if (t.description) lines.push(`DESCRIPTION:${icsText(t.description)}`);
    lines.push(`STATUS:${t.status === 'Cancelled' ? 'CANCELLED' : 'CONFIRMED'}`, 'TRANSP:TRANSPARENT', 'END:VEVENT');
  });
  lines.push('END:VCALENDAR');
  const blob = new Blob([lines.map(icsFold).join('\r\n') + '\r\n'], { type: 'text/calendar;charset=utf-8' });
  downloadBlob(blob, safeFileName(`${project.title}${strings.Svc_IcsFileSuffix}`));
}
