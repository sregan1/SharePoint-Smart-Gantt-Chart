import * as XLSX from 'xlsx';
import PptxGenJS from 'pptxgenjs';
import {
  addDays, differenceInCalendarDays, endOfMonth, format,
  isWeekend, startOfMonth, startOfWeek, addWeeks, addMonths, max, min, getISOWeek,
} from 'date-fns';
import {
  IProject, ITask, IGanttDisplaySettings, STATUS_COLORS, PRIORITY_COLORS,
  HEADER_THEME_COLORS, phaseColor,
} from '../models';

// ─── Excel export ─────────────────────────────────────────────────────────────

export function exportTasksToExcel(project: IProject, tasks: ITask[]): void {
  const fmt = (d: string): string => d ? format(new Date(d), 'MM/dd/yyyy') : '';

  const rows = tasks.map(t => ({
    'Task Name': t.title,
    'Phase': t.phase,
    'Start Date': fmt(t.startDate),
    'Due Date': fmt(t.dueDate),
    'Status': t.status,
    'Priority': t.priority,
    'Assigned To': t.assignedTo,
    'Assigned To (Email)': t.assignedToEmail,
    '% Complete': t.percentComplete,
    'Is Milestone': t.isMilestone ? 'Yes' : 'No',
    'Description': t.description,
    'Notes': t.notes,
  }));

  const ws = XLSX.utils.json_to_sheet(rows);

  // Auto-width
  const headers = Object.keys(rows[0] || {});
  ws['!cols'] = headers.map(h => ({
    wch: Math.max(h.length + 2, ...rows.map(r => String((r as Record<string, unknown>)[h] ?? '').length + 1)),
  }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Tasks');
  XLSX.writeFile(wb, `${project.title} - Tasks.xlsx`);
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

function parseDate(s: string): Date | null {
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

  tasks.filter(t => t.parentTaskId).forEach(t => {
    if (!subMap.has(t.parentTaskId!)) subMap.set(t.parentTaskId!, []);
    subMap.get(t.parentTaskId!)!.push(t);
  });

  tasks.filter(t => !t.parentTaskId).forEach(t => {
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
  if (task.color) return task.color;
  if (settings.colorBy === 'priority') return PRIORITY_COLORS[task.priority] || '#0078D4';
  if (settings.colorBy === 'phase' && task.phase) return phaseColor(task.phase);
  return STATUS_COLORS[task.status] || '#0078D4';
}

export function renderGanttSVG(
  project: IProject,
  tasks: ITask[],
  settings: IGanttDisplaySettings
): string {
  const ROW_H = settings.rowHeight;
  const theme = HEADER_THEME_COLORS[settings.headerTheme];

  // Date range
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const allDates: Date[] = [addDays(today, -14)];
  tasks.forEach(t => {
    const s = parseDate(t.startDate); const e = parseDate(t.dueDate);
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
    const s = parseDate(t.startDate);
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
    const sd = parseDate(task.startDate); const ed = parseDate(task.dueDate);
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

    // Right panel row line
    bars.push(`<line x1="${LEFT_W}" y1="${y0 + ROW_H}" x2="${LEFT_W + timelineW}" y2="${y0 + ROW_H}" stroke="#F3F2F1" stroke-width="1"/>`);

    // Dependency arrows
    if (settings.showDependencies) {
      task.dependencies.forEach(depId => {
        const depIdx = taskIndexMap.get(depId);
        if (depIdx === undefined) return;
        const depTask = visRows[depIdx]?.task;
        if (!depTask) return;
        const depEnd = parseDate(depTask.dueDate);
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
  <rect width="${totalW}" height="${TITLE_H}" fill="${project.color}"/>
  <circle cx="24" cy="${TITLE_H / 2}" r="7" fill="white" opacity="0.25"/>
  <text x="38" y="${TITLE_H / 2 + 6}" font-size="17" font-weight="700" fill="white">${escXml(project.title)}</text>
  <text x="${totalW - 12}" y="${TITLE_H / 2 + 5}" font-size="11" fill="rgba(255,255,255,0.7)" text-anchor="end">Exported ${format(today, 'MMM d, yyyy')}</text>

  <!-- ── Left panel header ─────────────────────────────────────────── -->
  <rect y="${TITLE_H}" width="${LEFT_W}" height="${HEADER_H}" fill="${theme.bg}"/>
  <text x="16" y="${TITLE_H + HEADER_H / 2 + 4}" font-size="11" font-weight="600" fill="${theme.subtext}" letter-spacing="0.5">TASK NAME</text>

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

export function downloadSVG(svgString: string, filename: string): void {
  const blob = new Blob([svgString], { type: 'image/svg+xml' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export function downloadPNG(svgString: string, filename: string, scale: number = 2): Promise<void> {
  return svgToCanvas(svgString, scale).then(canvas => new Promise<void>((resolve) => {
    canvas.toBlob(pngBlob => {
      if (!pngBlob) { resolve(); return; }
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

// Shared SVG→canvas helper used by both downloadPNG and svgToPngDataUrl.
function svgToCanvas(svgString: string, scale: number): Promise<HTMLCanvasElement> {
  return new Promise((resolve, reject) => {
    const blob = new Blob([svgString], { type: 'image/svg+xml' });
    const svgUrl = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = (): void => {
      const canvas = document.createElement('canvas');
      canvas.width = img.width * scale;
      canvas.height = img.height * scale;
      const ctx = canvas.getContext('2d');
      if (!ctx) { URL.revokeObjectURL(svgUrl); resolve(canvas); return; }
      ctx.scale(scale, scale);
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, img.width, img.height);
      ctx.drawImage(img, 0, 0);
      URL.revokeObjectURL(svgUrl);
      resolve(canvas);
    };
    img.onerror = (): void => reject(new Error('SVG render failed'));
    img.src = svgUrl;
  });
}

function svgToPngDataUrl(svgString: string, scale: number = 2): Promise<string> {
  return svgToCanvas(svgString, scale).then(canvas => canvas.toDataURL('image/png'));
}

function hex(color: string): string {
  return color.startsWith('#') ? color.slice(1) : color;
}

export async function exportToPowerPoint(
  project: IProject,
  tasks: ITask[],
  settings: IGanttDisplaySettings
): Promise<void> {
  const today = new Date();
  const fmt = (d: string): string => d ? format(new Date(d), 'MMM d, yyyy') : '—';
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

  cover.addText('PROJECT REPORT', {
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

  cover.addText(`Status: ${project.status}`, {
    x: 0.65, y: 3.1, w: 6, h: 0.45,
    fontSize: 16,
    color: 'FFFFFF',
    transparency: 15,
  });

  const dateRange = (project.startDate || project.dueDate)
    ? `${fmt(project.startDate)}  →  ${fmt(project.dueDate)}`
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
    cover.addText(`Project Manager: ${project.projectManager}`, {
      x: 0.65, y: 5.8, w: 8, h: 0.35,
      fontSize: 12,
      color: '605E5C',
    });
  }

  cover.addText(`Generated ${format(today, 'MMMM d, yyyy')}`, {
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
  summary.addText('Project Summary', {
    x: 0.4, y: 0, w: 9, h: 0.85,
    fontSize: 22, color: 'FFFFFF', bold: true, valign: 'middle',
  });
  summary.addText(project.title, {
    x: 0, y: 0, w: 13.0, h: 0.85,
    fontSize: 13, color: 'FFFFFF', align: 'right', valign: 'middle', transparency: 35,
  });

  // Stat boxes
  const statItems = [
    { label: 'Total Tasks',  value: totalCount,                   bg: 'F3F2F1', fg: '323130' },
    { label: 'Completed',    value: byStatus['Completed'],        bg: 'F1FAF1', fg: '107C10' },
    { label: 'In Progress',  value: byStatus['In Progress'],      bg: 'EFF6FC', fg: '0078D4' },
    { label: 'On Hold',      value: byStatus['On Hold'],          bg: 'FFF4EC', fg: 'CA5010' },
    { label: 'Not Started',  value: byStatus['Not Started'],      bg: 'F3F2F1', fg: '605E5C' },
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
  summary.addText(`Overall Progress: ${overallPct}%`, {
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
  summary.addText(`${overallPct}%`, {
    x: 12.85, y: BAR_Y, w: 0.8, h: 0.35,
    fontSize: 13, bold: true, color: projectColor, align: 'right',
  });

  // Status breakdown (two columns)
  const TBL_Y = 3.75;
  summary.addText('Status Breakdown', {
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
    summary.addText(`${s.label}: ${cnt}  (${pct}%)`, {
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
  gantt.addText('Gantt Timeline', {
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
  activity.addText('Summary & Recent Activity', {
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

  activity.addText('PROJECT OVERVIEW', {
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
    { label: 'Status', value: project.status },
    ...(project.startDate ? [{ label: 'Start Date', value: fmt(project.startDate) }] : []),
    ...(project.dueDate   ? [{ label: 'Due Date',   value: fmt(project.dueDate)   }] : []),
    ...(project.projectManager ? [{ label: 'Project Manager', value: project.projectManager }] : []),
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

  activity.addText('PAST 7 DAYS', {
    x: RX, y: 1.05, w: RW, h: 0.3,
    fontSize: 10, bold: true, color: '605E5C', charSpacing: 1.5,
  });

  const weekAgo = addDays(today, -7);
  const recentTasks = tasks.filter(t => {
    const mod = t.modified ? parseDate(t.modified) : null;
    return mod !== null && mod >= weekAgo;
  });

  const completedThisWeek = recentTasks.filter(t => t.status === 'Completed');
  const updatedThisWeek   = recentTasks.filter(t => t.status !== 'Completed');

  let ry = 1.45;
  const MAX_PER_SECTION = 4;

  if (recentTasks.length === 0) {
    activity.addText('No task activity recorded in the past 7 days.', {
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
      activity.addText(`Completed this week  (${completedThisWeek.length})`, {
        x: RX + 0.15, y: ry, w: RW - 0.3, h: 0.33,
        fontSize: 11, bold: true, color: '107C10', valign: 'middle',
      });
      ry += 0.4;

      completedThisWeek.slice(0, MAX_PER_SECTION).forEach(t => {
        const isNew = !!(t.created && parseDate(t.created) !== null && parseDate(t.created)! >= weekAgo);
        const titleRuns = isNew
          ? [{ text: 'NEW  ', options: { color: '107C10', bold: true, fontSize: 9 } },
             { text: t.title, options: { color: '323130', fontSize: 12 } }]
          : [{ text: t.title, options: { color: '323130', fontSize: 12 } }];

        activity.addShape(pptx.ShapeType.ellipse, {
          x: RX + 0.1, y: ry + 0.1, w: 0.15, h: 0.15,
          fill: { color: '107C10' }, line: { color: '107C10', width: 0 },
        });
        activity.addText(titleRuns, {
          x: RX + 0.34, y: ry, w: RW - 0.34, h: 0.26, valign: 'middle',
        });
        activity.addText(`Completed  •  ${t.percentComplete}%`, {
          x: RX + 0.34, y: ry + 0.27, w: RW - 0.34, h: 0.18,
          fontSize: 9, color: '8A8886',
        });
        ry += 0.46;
      });

      if (completedThisWeek.length > MAX_PER_SECTION) {
        activity.addText(`+ ${completedThisWeek.length - MAX_PER_SECTION} more`, {
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
      activity.addText(`In progress / updated  (${updatedThisWeek.length})`, {
        x: RX + 0.15, y: ry, w: RW - 0.3, h: 0.33,
        fontSize: 11, bold: true, color: '0078D4', valign: 'middle',
      });
      ry += 0.4;

      updatedThisWeek.slice(0, MAX_PER_SECTION).forEach(t => {
        const isNew = !!(t.created && parseDate(t.created) !== null && parseDate(t.created)! >= weekAgo);
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
        activity.addText(`${t.status}  •  ${t.percentComplete}%`, {
          x: RX + 0.34, y: ry + 0.27, w: RW - 0.34, h: 0.18,
          fontSize: 9, color: '8A8886',
        });
        ry += 0.46;
      });

      if (updatedThisWeek.length > MAX_PER_SECTION) {
        activity.addText(`+ ${updatedThisWeek.length - MAX_PER_SECTION} more`, {
          x: RX + 0.34, y: ry, w: RW, h: 0.3,
          fontSize: 11, color: '8A8886', italic: true,
        });
      }
    }
  }

  await pptx.writeFile({ fileName: `${project.title} - Project Report.pptx` });
}
