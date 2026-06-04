'use strict';
// Generates mockup screenshots of the Smart Gantt Chart web part.
// Run with: node docs/generate-screenshots.js

const puppeteer = require('puppeteer-core');
const path = require('path');
const fs = require('fs');

const CHROME = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const OUT = path.join(__dirname, 'screenshots');
if (!fs.existsSync(OUT)) fs.mkdirSync(OUT, { recursive: true });

// ── Dummy data ─────────────────────────────────────────────────────────────
const PROJECT = { title: 'Website Redesign', color: '#0078D4', status: 'Active' };
const TODAY = new Date('2026-06-04');

const TASKS = [
  { id:1, phase:'Discovery', title:'Requirements Gathering', status:'Completed',   pct:100, assignee:'Sarah M.', start:'2026-05-15', end:'2026-05-27', priority:'High'   },
  { id:2, phase:'Discovery', title:'Stakeholder Interviews',  status:'Completed',   pct:100, assignee:'John S.',  start:'2026-05-18', end:'2026-05-29', priority:'Medium' },
  { id:3, phase:'Discovery', title:'Current State Analysis',  status:'In Progress', pct:75,  assignee:'Sarah M.', start:'2026-05-28', end:'2026-06-06', priority:'High'   },
  { id:4, phase:'Design',    title:'UX Wireframes',           status:'In Progress', pct:40,  assignee:'Amy K.',   start:'2026-06-03', end:'2026-06-17', priority:'Medium' },
  { id:5, phase:'Design',    title:'Visual Design',           status:'Not Started', pct:0,   assignee:'Amy K.',   start:'2026-06-10', end:'2026-06-24', priority:'Medium' },
  { id:6, phase:'Design',    title:'Design Review',           status:'Not Started', pct:0,   assignee:'',         start:'2026-06-27', end:'2026-06-27', milestone:true,   priority:'High'   },
  { id:7, phase:'Development', title:'Frontend Build',        status:'Not Started', pct:0,   assignee:'Mike R.',  start:'2026-06-26', end:'2026-07-17', priority:'Medium' },
  { id:8, phase:'Development', title:'API Integration',       status:'Not Started', pct:0,   assignee:'Dave C.',  start:'2026-07-08', end:'2026-07-23', priority:'High'   },
  { id:9, phase:'Development', title:'QA Testing',            status:'Not Started', pct:0,   assignee:'Sarah M.', start:'2026-07-21', end:'2026-08-04', priority:'Medium' },
  { id:10, phase:'Development', title:'Go Live',              status:'Not Started', pct:0,   assignee:'',         start:'2026-08-07', end:'2026-08-07', milestone:true,   priority:'Critical' },
];

// ── Colors ─────────────────────────────────────────────────────────────────
const STATUS_COLOR = {
  'Not Started': '#8B929A', 'In Progress': '#0078D4',
  'Completed': '#107C10',   'On Hold': '#CA5010', 'Cancelled': '#D13438',
};
const STATUS_BG = {
  'Not Started': '#F3F2F1', 'In Progress': '#EFF6FC',
  'Completed': '#F1FAF1',   'On Hold': '#FFF4EC',  'Cancelled': '#FDF3F4',
};
const PRIORITY_COLOR = { Critical:'#D13438', High:'#CA5010', Medium:'#0078D4', Low:'#107C10' };

// ── Gantt SVG ──────────────────────────────────────────────────────────────
function ganttSVG() {
  const D = 8; // px per day
  const LW = 300, TH = 48, HH = 56, RH = 40, BH = 24;
  const range0 = new Date('2026-05-01');
  const range1 = new Date('2026-08-31');
  const days = (d) => Math.round((d - range0) / 86400000);
  const totalDays = days(range1) + 1;
  const TW = totalDays * D;
  const W  = LW + TW;

  // rows: phases + tasks
  const rows = [];
  let lastPhase = null;
  for (const t of TASKS) {
    if (t.phase !== lastPhase) { rows.push({ type:'phase', phase: t.phase }); lastPhase = t.phase; }
    rows.push({ type:'task', task: t });
  }
  const bodyH = rows.length * RH + 20;
  const H = TH + HH + bodyH;

  // month bands
  const months = [
    { label:'May 2026', x:0,   w:31*D },
    { label:'Jun 2026', x:31*D, w:30*D },
    { label:'Jul 2026', x:61*D, w:31*D },
    { label:'Aug 2026', x:92*D, w:31*D },
  ];

  // week bands (Mondays; May 1 = Friday, so Mon Apr 27 = day -4)
  const weekBands = [];
  const todayDate = TODAY;
  const thisMonday = new Date(todayDate);
  thisMonday.setDate(todayDate.getDate() - ((todayDate.getDay() + 6) % 7));

  let wStart = new Date('2026-04-27'); // Mon before May 1
  let wNum = 18;
  while (wStart <= range1) {
    const wx0 = Math.max(0, days(wStart));
    const wx1 = Math.min(totalDays - 1, days(wStart) + 6);
    if (wx1 >= 0) {
      const isCurrent = wStart.getTime() === thisMonday.getTime();
      weekBands.push({ x: wx0 * D, w: (wx1 - wx0 + 1) * D, label: `W${wNum}`, current: isCurrent });
    }
    wStart = new Date(wStart); wStart.setDate(wStart.getDate() + 7);
    wNum++;
  }

  // weekend rects
  const weekendRects = [];
  for (let d = 0; d < totalDays; d++) {
    const dt = new Date(range0); dt.setDate(range0.getDate() + d);
    const dow = dt.getDay();
    if (dow === 0 || dow === 6) {
      weekendRects.push(`<rect x="${LW + d*D}" y="${TH+HH}" width="${D}" height="${bodyH}" fill="#F8F7F6"/>`);
    }
  }

  // rows
  const bars = []; const todayX = LW + days(TODAY)*D + D/2;
  rows.forEach((row, i) => {
    const y0 = TH + HH + i * RH;
    if (row.type === 'phase') {
      bars.push(
        `<rect x="0" y="${y0}" width="${LW}" height="${RH}" fill="#F3F2F1"/>`,
        `<rect x="${LW}" y="${y0}" width="${TW}" height="${RH}" fill="#F8F7F6"/>`,
        `<text x="14" y="${y0+RH/2+4}" font-size="11" font-weight="700" fill="#605E5C" font-family="Segoe UI,sans-serif" letter-spacing="0.5">${row.phase.toUpperCase()}</text>`,
      );
      return;
    }
    const t = row.task;
    const sc = STATUS_COLOR[t.status] || '#0078D4';
    bars.push(
      `<rect x="0" y="${y0}" width="${LW}" height="${RH}" fill="${i%2===0?'#FFFFFF':'#FAFAFA'}"/>`,
      `<line x1="0" y1="${y0+RH}" x2="${LW}" y2="${y0+RH}" stroke="#F3F2F1" stroke-width="1"/>`,
      `<circle cx="10" cy="${y0+RH/2}" r="4" fill="${sc}"/>`,
      `<text x="22" y="${y0+RH/2+4}" font-size="12" fill="#323130" font-family="Segoe UI,sans-serif">${t.title.substring(0,36)}</text>`,
    );
    const s = new Date(t.start), e = new Date(t.end);
    const bx = LW + days(s)*D;
    const bw = Math.max(8, (days(e)-days(s)+1)*D);
    const by = y0 + (RH-BH)/2;
    const pw = bw * t.pct/100;

    if (t.milestone) {
      const mx = LW + days(s)*D + D/2, my = y0+RH/2, ms=9;
      bars.push(`<polygon points="${mx},${my-ms} ${mx+ms},${my} ${mx},${my+ms} ${mx-ms},${my}" fill="${sc}" stroke="white" stroke-width="1.5"/>`);
    } else {
      bars.push(
        `<rect x="${bx}" y="${by}" width="${bw}" height="${BH}" rx="4" fill="${sc}28"/>`,
        pw>0 ? `<rect x="${bx}" y="${by}" width="${pw}" height="${BH}" rx="4" fill="${sc}"/>` : '',
      );
      if (t.pct > 0 && bw > 40) bars.push(
        `<text x="${bx+6}" y="${by+BH/2+4}" font-size="10" font-weight="700" fill="${pw>bw*.45?'#fff':sc}" font-family="Segoe UI,sans-serif">${t.pct}%</text>`
      );
    }
    bars.push(`<line x1="${LW}" y1="${y0+RH}" x2="${LW+TW}" y2="${y0+RH}" stroke="#F3F2F1" stroke-width="1"/>`);
  });

  const theme = { bg:'#1B1B3A', text:'rgba(255,255,255,0.9)', sub:'rgba(255,255,255,0.55)' };

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" style="background:white;font-family:'Segoe UI',Arial,sans-serif">
  <!-- title bar -->
  <rect width="${W}" height="${TH}" fill="${PROJECT.color}"/>
  <circle cx="24" cy="${TH/2}" r="7" fill="white" opacity="0.25"/>
  <text x="38" y="${TH/2+6}" font-size="17" font-weight="700" fill="white" font-family="Segoe UI,sans-serif">${PROJECT.title}</text>
  <text x="${W-12}" y="${TH/2+5}" font-size="11" fill="rgba(255,255,255,0.7)" text-anchor="end" font-family="Segoe UI,sans-serif">Jun 4, 2026</text>

  <!-- header bg -->
  <rect y="${TH}" width="${LW}" height="${HH}" fill="${theme.bg}"/>
  <rect x="${LW}" y="${TH}" width="${TW}" height="${HH}" fill="${theme.bg}"/>
  <text x="16" y="${TH+HH/2+4}" font-size="11" font-weight="600" fill="${theme.sub}" letter-spacing="0.5" font-family="Segoe UI,sans-serif">TASK NAME</text>

  <!-- month bands -->
  ${months.map(m => `
    <line x1="${LW+m.x}" y1="${TH}" x2="${LW+m.x}" y2="${TH+28}" stroke="rgba(255,255,255,0.15)" stroke-width="1"/>
    <text x="${LW+m.x+6}" y="${TH+18}" font-size="12" font-weight="600" fill="${theme.text}" font-family="Segoe UI,sans-serif">${m.label}</text>
  `).join('')}

  <!-- week bands -->
  ${weekBands.map(w => `
    ${w.current ? `<rect x="${LW+w.x}" y="${TH+28}" width="${w.w}" height="28" fill="rgba(255,215,0,0.15)"/>` : ''}
    <line x1="${LW+w.x}" y1="${TH+28}" x2="${LW+w.x}" y2="${TH+HH}" stroke="rgba(255,255,255,0.1)" stroke-width="1"/>
    <text x="${LW+w.x+3}" y="${TH+HH-9}" font-size="10" fill="${w.current?'#FFD700':theme.sub}" font-weight="${w.current?'700':'400'}" font-family="Segoe UI,sans-serif">${w.label}</text>
  `).join('')}

  <!-- body -->
  <rect y="${TH+HH}" width="${W}" height="${bodyH}" fill="white"/>
  ${weekendRects.join('\n  ')}
  ${bars.join('\n  ')}

  <!-- divider -->
  <line x1="${LW}" y1="${TH}" x2="${LW}" y2="${H}" stroke="#EDEBE9" stroke-width="2"/>

  <!-- today -->
  <line x1="${todayX}" y1="${TH+HH}" x2="${todayX}" y2="${H}" stroke="#D13438" stroke-width="2" stroke-dasharray="4,3"/>
  <circle cx="${todayX}" cy="${TH+HH}" r="4" fill="#D13438"/>
</svg>`;
}

// ── Shared CSS ─────────────────────────────────────────────────────────────
const CSS = `
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', Arial, sans-serif; background: #FAF9F8; color: #323130; font-size: 14px; }
.webpart { background: #fff; border-radius: 4px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); overflow: hidden; }

/* Toolbar */
.toolbar { background: #fff; border-bottom: 1px solid #EDEBE9; padding: 0 16px; }
.row1, .row2 { display: flex; align-items: center; gap: 8px; height: 44px; }
.row2 { border-top: 1px solid #F3F2F1; }
.project-selector { display:flex; align-items:center; gap:8px; padding:6px 10px; border:1px solid #EDEBE9; border-radius:4px; cursor:pointer; min-width:200px; background:#fff; }
.project-dot { width:10px; height:10px; border-radius:50%; background:${PROJECT.color}; flex-shrink:0; }
.project-name { font-weight:600; font-size:13px; color:#323130; }
.chevron { color:#605E5C; font-size:11px; margin-left:auto; }
.divider { width:1px; height:24px; background:#EDEBE9; margin:0 4px; }
.btn { display:inline-flex; align-items:center; gap:5px; padding:5px 12px; border-radius:4px; font-size:13px; cursor:pointer; border:1px solid transparent; font-family:inherit; }
.btn-primary { background:${PROJECT.color}; color:#fff; border-color:${PROJECT.color}; font-weight:600; }
.btn-secondary { background:#fff; color:#323130; border-color:#EDEBE9; }
.btn-ghost { background:transparent; color:#323130; border:none; padding:5px 8px; }
.zoom-group { display:flex; }
.zoom-btn { padding:4px 10px; font-size:12px; background:#fff; border:1px solid #EDEBE9; cursor:pointer; font-family:inherit; color:#323130; }
.zoom-btn:first-child { border-radius:4px 0 0 4px; }
.zoom-btn:last-child  { border-radius:0 4px 4px 0; }
.zoom-btn.active { background:${PROJECT.color}; color:#fff; border-color:${PROJECT.color}; font-weight:600; }
.view-switcher { display:flex; border:1px solid #EDEBE9; border-radius:4px; overflow:hidden; }
.view-btn { display:flex; align-items:center; gap:4px; padding:5px 12px; font-size:12px; background:#fff; border:none; border-right:1px solid #EDEBE9; cursor:pointer; font-family:inherit; color:#605E5C; }
.view-btn:last-child { border-right:none; }
.view-btn.active { background:#EFF6FC; color:${PROJECT.color}; font-weight:600; }
.settings-btn { padding:5px 10px; font-size:12px; background:#fff; border:1px solid #EDEBE9; border-radius:4px; cursor:pointer; font-family:inherit; color:#323130; }
.icon-btn { padding:5px 10px; font-size:16px; background:#fff; border:1px solid #EDEBE9; border-radius:4px; cursor:pointer; font-family:inherit; color:#323130; line-height:1; }
.row2-left { display:flex; align-items:center; gap:8px; flex:1; }
.row2-right { display:flex; align-items:center; gap:6px; margin-left:auto; }
.today-btn { padding:5px 10px; font-size:12px; background:#fff; border:1px solid #EDEBE9; border-radius:4px; cursor:pointer; color:#323130; font-family:inherit; }
`;

// ── Toolbar HTML ───────────────────────────────────────────────────────────
function toolbar(view, extra = '') {
  return `<div class="toolbar">
  <div class="row1">
    <div class="project-selector">
      <div class="project-dot"></div>
      <span class="project-name">${PROJECT.title}</span>
      <span class="chevron">▾</span>
    </div>
    <div class="divider"></div>
    <button class="btn btn-primary">+ Add Task</button>
    <button class="btn btn-secondary">Edit Project</button>
    ${extra}
  </div>
  <div class="row2">
    <div class="row2-left">
      ${view === 'gantt' ? `
        <button class="today-btn">◉ Today</button>
        <div class="zoom-group">
          <button class="zoom-btn">Day</button>
          <button class="zoom-btn active">Week</button>
          <button class="zoom-btn">Month</button>
          <button class="zoom-btn">Quarter</button>
        </div>
        <div class="divider"></div>
      ` : ''}
    </div>
    <div class="row2-right">
      <div class="view-switcher">
        <button class="view-btn ${view==='list'?'active':''}">☰ List</button>
        <button class="view-btn ${view==='gantt'?'active':''}">▬ Gantt</button>
        <button class="view-btn ${view==='kanban'?'active':''}">⬜ Kanban</button>
      </div>
      ${view==='gantt' ? '<button class="settings-btn">⚙ Display</button>' : ''}
      <button class="icon-btn">⋯</button>
    </div>
  </div>
</div>`;
}

// ── Wrapper ────────────────────────────────────────────────────────────────
function page(body, extraCss = '') {
  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>${CSS}${extraCss}</style></head><body>
<div style="padding:20px;background:#FAF9F8;min-height:100vh;">
<div class="webpart">${body}</div>
</div></body></html>`;
}

// ── 1. Gantt view ──────────────────────────────────────────────────────────
function ganttPage() {
  const svg = ganttSVG();
  return page(`
    ${toolbar('gantt')}
    <div style="overflow:auto; background:#fff;">
      ${svg}
    </div>
  `);
}

// ── 2. List view ───────────────────────────────────────────────────────────
function listPage() {
  const fmtDate = d => new Date(d).toLocaleDateString('en-US', {month:'short',day:'numeric',year:'numeric'});
  const isOverdue = d => new Date(d) < TODAY;

  let rows = '';
  let lastPhase = null;
  for (const t of TASKS) {
    if (t.phase !== lastPhase) {
      rows += `<tr style="background:#F3F2F1;"><td colspan="7" style="padding:6px 12px;font-size:11px;font-weight:700;color:#605E5C;letter-spacing:0.5px;">${t.phase.toUpperCase()}</td></tr>`;
      lastPhase = t.phase;
    }
    const sc = STATUS_COLOR[t.status], sbg = STATUS_BG[t.status];
    const pc = PRIORITY_COLOR[t.priority];
    const due = isOverdue(t.end) && t.status !== 'Completed';
    rows += `<tr style="border-bottom:1px solid #F3F2F1;">
      <td style="padding:10px 12px;display:flex;align-items:center;gap:8px;">
        <span style="width:8px;height:8px;border-radius:50%;background:${sc};display:inline-block;flex-shrink:0;"></span>
        <span style="font-size:13px;color:#323130;">${t.milestone?'◆ ':''}${t.title}</span>
      </td>
      <td style="padding:10px 12px;"><span style="padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600;background:${sbg};color:${sc};">${t.status}</span></td>
      <td style="padding:10px 12px;"><span style="padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600;background:${sbg};color:${pc};">${t.priority}</span></td>
      <td style="padding:10px 12px;font-size:12px;color:#605E5C;">${fmtDate(t.start)}</td>
      <td style="padding:10px 12px;font-size:12px;color:${due?'#D13438':'#605E5C'};">${fmtDate(t.end)}${due?' ⚠':''}</td>
      <td style="padding:10px 12px;font-size:12px;color:#323130;">${t.assignee||'—'}</td>
      <td style="padding:10px 12px;min-width:120px;">
        <div style="display:flex;align-items:center;gap:6px;">
          <div style="flex:1;height:6px;background:#EDEBE9;border-radius:3px;overflow:hidden;">
            <div style="width:${t.pct}%;height:100%;background:${sc};border-radius:3px;"></div>
          </div>
          <span style="font-size:11px;color:#605E5C;width:28px;text-align:right;">${t.pct}%</span>
        </div>
      </td>
    </tr>`;
  }

  return page(`
    ${toolbar('list')}
    <div style="overflow:auto;background:#fff;">
      <table style="width:100%;border-collapse:collapse;">
        <thead>
          <tr style="background:#F3F2F1;border-bottom:1px solid #EDEBE9;">
            <th style="padding:10px 12px;font-size:11px;font-weight:600;color:#605E5C;text-align:left;min-width:250px;">TASK NAME ↑</th>
            <th style="padding:10px 12px;font-size:11px;font-weight:600;color:#605E5C;text-align:left;">STATUS</th>
            <th style="padding:10px 12px;font-size:11px;font-weight:600;color:#605E5C;text-align:left;">PRIORITY</th>
            <th style="padding:10px 12px;font-size:11px;font-weight:600;color:#605E5C;text-align:left;">START</th>
            <th style="padding:10px 12px;font-size:11px;font-weight:600;color:#605E5C;text-align:left;">DUE</th>
            <th style="padding:10px 12px;font-size:11px;font-weight:600;color:#605E5C;text-align:left;">ASSIGNED TO</th>
            <th style="padding:10px 12px;font-size:11px;font-weight:600;color:#605E5C;text-align:left;min-width:120px;">PROGRESS</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `);
}

// ── 3. Kanban view ─────────────────────────────────────────────────────────
function kanbanPage() {
  const cols = ['Not Started','In Progress','On Hold','Completed','Cancelled'];
  const fmtDate = d => new Date(d).toLocaleDateString('en-US', {month:'short',day:'numeric'});

  function card(t) {
    const sc = STATUS_COLOR[t.status], pc = PRIORITY_COLOR[t.priority];
    const initials = t.assignee ? t.assignee.split(' ').map(w=>w[0]).join('').substring(0,2) : null;
    return `<div style="background:#fff;border:1px solid #EDEBE9;border-radius:6px;padding:10px;margin-bottom:8px;box-shadow:0 1px 3px rgba(0,0,0,0.04);">
      <div style="display:flex;align-items:flex-start;gap:6px;margin-bottom:6px;">
        <div style="width:6px;height:6px;border-radius:50%;background:${pc};flex-shrink:0;margin-top:5px;"></div>
        <span style="font-size:13px;font-weight:600;color:#323130;line-height:1.3;">${t.milestone?'◆ ':''}${t.title}</span>
      </div>
      <div style="display:flex;gap:4px;flex-wrap:wrap;margin-bottom:6px;">
        <span style="padding:1px 6px;border-radius:10px;font-size:10px;font-weight:600;background:${STATUS_BG[t.status]};color:${sc};">${t.status}</span>
        <span style="padding:1px 6px;border-radius:10px;font-size:10px;background:#F3F2F1;color:#605E5C;">${t.phase}</span>
      </div>
      ${!t.milestone ? `<div style="font-size:11px;color:#605E5C;margin-bottom:6px;">${fmtDate(t.start)} → ${fmtDate(t.end)}</div>` : ''}
      <div style="display:flex;align-items:center;gap:6px;">
        <div style="flex:1;height:4px;background:#EDEBE9;border-radius:2px;overflow:hidden;">
          <div style="width:${t.pct}%;height:100%;background:${sc};border-radius:2px;"></div>
        </div>
        ${initials ? `<div style="width:20px;height:20px;border-radius:50%;background:${sc};color:#fff;font-size:9px;font-weight:700;display:flex;align-items:center;justify-content:center;">${initials}</div>` : ''}
      </div>
    </div>`;
  }

  const colHtml = cols.map(col => {
    const colTasks = TASKS.filter(t => t.status === col);
    return `<div style="flex:1;min-width:0;background:#F3F2F1;border-radius:6px;padding:10px;display:flex;flex-direction:column;">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;">
        <span style="font-size:12px;font-weight:700;color:#323130;">${col}</span>
        <span style="background:#EDEBE9;color:#605E5C;font-size:11px;font-weight:600;padding:1px 7px;border-radius:10px;">${colTasks.length}</span>
      </div>
      ${colTasks.map(card).join('')}
      <div style="margin-top:4px;text-align:center;padding:6px;font-size:12px;color:#8A8886;cursor:pointer;">+ Add Task</div>
    </div>`;
  }).join('');

  return page(`
    ${toolbar('kanban')}
    <div style="padding:16px;display:flex;gap:12px;background:#FAF9F8;min-height:600px;">
      ${colHtml}
    </div>
  `);
}

// ── 4. Task panel ──────────────────────────────────────────────────────────
function taskPanelPage() {
  const t = TASKS[3]; // UX Wireframes
  const svg = ganttSVG();

  return page(`
    ${toolbar('gantt')}
    <div style="display:flex;position:relative;overflow:hidden;">
      <div style="flex:1;overflow:auto;">${svg}</div>
      <div style="width:420px;flex-shrink:0;border-left:1px solid #EDEBE9;background:#fff;display:flex;flex-direction:column;height:740px;overflow:auto;">
        <div style="padding:16px 20px;border-bottom:1px solid #EDEBE9;display:flex;align-items:center;justify-content:space-between;">
          <span style="font-size:16px;font-weight:700;color:#323130;">Edit Task</span>
          <button style="border:none;background:none;font-size:18px;color:#605E5C;cursor:pointer;">✕</button>
        </div>
        <!-- Tabs -->
        <div style="display:flex;border-bottom:1px solid #EDEBE9;">
          <div style="padding:10px 20px;font-size:13px;font-weight:600;color:${PROJECT.color};border-bottom:2px solid ${PROJECT.color};cursor:pointer;">Basic</div>
          <div style="padding:10px 20px;font-size:13px;color:#605E5C;cursor:pointer;">Details</div>
          <div style="padding:10px 20px;font-size:13px;color:#605E5C;cursor:pointer;">Links</div>
        </div>
        <!-- Fields -->
        <div style="padding:20px;display:flex;flex-direction:column;gap:16px;flex:1;">
          <div>
            <label style="display:block;font-size:12px;font-weight:600;color:#323130;margin-bottom:4px;">Task Name <span style="color:#D13438;">*</span></label>
            <input style="width:100%;padding:7px 10px;border:1px solid #EDEBE9;border-radius:4px;font-size:13px;font-family:inherit;color:#323130;" value="${t.title}"/>
          </div>
          <div style="display:flex;gap:12px;">
            <div style="flex:1;">
              <label style="display:block;font-size:12px;font-weight:600;color:#323130;margin-bottom:4px;">Start Date</label>
              <input type="text" style="width:100%;padding:7px 10px;border:1px solid #EDEBE9;border-radius:4px;font-size:13px;font-family:inherit;color:#323130;" value="Jun 3, 2026"/>
            </div>
            <div style="flex:1;">
              <label style="display:block;font-size:12px;font-weight:600;color:#323130;margin-bottom:4px;">Due Date</label>
              <input type="text" style="width:100%;padding:7px 10px;border:1px solid #EDEBE9;border-radius:4px;font-size:13px;font-family:inherit;color:#323130;" value="Jun 17, 2026"/>
            </div>
          </div>
          <div style="display:flex;gap:12px;">
            <div style="flex:1;">
              <label style="display:block;font-size:12px;font-weight:600;color:#323130;margin-bottom:4px;">Status</label>
              <select style="width:100%;padding:7px 10px;border:1px solid #EDEBE9;border-radius:4px;font-size:13px;font-family:inherit;color:#323130;">
                <option>Not Started</option>
                <option selected>In Progress</option>
                <option>Completed</option>
                <option>On Hold</option>
                <option>Cancelled</option>
              </select>
            </div>
            <div style="flex:1;">
              <label style="display:block;font-size:12px;font-weight:600;color:#323130;margin-bottom:4px;">Priority</label>
              <select style="width:100%;padding:7px 10px;border:1px solid #EDEBE9;border-radius:4px;font-size:13px;font-family:inherit;color:#323130;">
                <option>Critical</option>
                <option>High</option>
                <option selected>Medium</option>
                <option>Low</option>
              </select>
            </div>
          </div>
          <div>
            <label style="display:block;font-size:12px;font-weight:600;color:#323130;margin-bottom:4px;">% Complete</label>
            <div style="display:flex;align-items:center;gap:10px;">
              <input type="range" min="0" max="100" value="40" style="flex:1;accent-color:${PROJECT.color};"/>
              <span style="font-size:13px;font-weight:600;color:${PROJECT.color};width:36px;text-align:right;">40%</span>
            </div>
          </div>
          <div>
            <label style="display:block;font-size:12px;font-weight:600;color:#323130;margin-bottom:4px;">Assigned To</label>
            <input style="width:100%;padding:7px 10px;border:1px solid #EDEBE9;border-radius:4px;font-size:13px;font-family:inherit;color:#323130;" value="${t.assignee}"/>
          </div>
          <div>
            <label style="display:block;font-size:12px;font-weight:600;color:#323130;margin-bottom:4px;">Description</label>
            <textarea style="width:100%;padding:7px 10px;border:1px solid #EDEBE9;border-radius:4px;font-size:13px;font-family:inherit;color:#323130;resize:none;height:72px;">Create low-fidelity wireframes for all key screens including homepage, product pages, and checkout flow.</textarea>
          </div>
        </div>
        <div style="padding:16px 20px;border-top:1px solid #EDEBE9;display:flex;gap:8px;justify-content:flex-end;">
          <button class="btn btn-secondary">Cancel</button>
          <button class="btn btn-primary">Save</button>
        </div>
      </div>
    </div>
  `);
}

// ── 5. Display settings ────────────────────────────────────────────────────
function displaySettingsPage() {
  const svg = ganttSVG();
  return page(`
    ${toolbar('gantt')}
    <div style="display:flex;position:relative;">
      <div style="flex:1;overflow:auto;">${svg}</div>
      <div style="width:300px;flex-shrink:0;border-left:1px solid #EDEBE9;background:#fff;height:740px;overflow:auto;padding:16px;">
        <div style="font-size:15px;font-weight:700;color:#323130;margin-bottom:16px;padding-bottom:12px;border-bottom:1px solid #EDEBE9;">Display Settings</div>

        <div style="margin-bottom:16px;">
          <div style="font-size:12px;font-weight:600;color:#323130;margin-bottom:8px;">Color Coding</div>
          ${['By Status','By Priority','By Phase'].map((o,i)=>`
          <label style="display:flex;align-items:center;gap:8px;padding:4px 0;cursor:pointer;">
            <input type="radio" name="color" ${i===0?'checked':''} style="accent-color:${PROJECT.color};"/> <span style="font-size:13px;color:#323130;">${o}</span>
          </label>`).join('')}
        </div>

        <div style="margin-bottom:16px;">
          <div style="font-size:12px;font-weight:600;color:#323130;margin-bottom:8px;">Header Theme</div>
          <div style="display:flex;gap:6px;">
            ${[['#1B1B3A','Dark'],['#1565C0','Navy'],['#00695C','Teal'],['#4527A0','Purple'],['#F3F2F1','Light']].map(([c,l],i)=>`
            <div style="width:32px;height:28px;background:${c};border-radius:4px;border:${i===0?`2px solid ${PROJECT.color}`:'2px solid transparent'};cursor:pointer;display:flex;align-items:center;justify-content:center;">
              ${i===0?`<span style="color:white;font-size:10px;">✓</span>`:''}
            </div>`).join('')}
          </div>
        </div>

        <div style="margin-bottom:16px;">
          <div style="font-size:12px;font-weight:600;color:#323130;margin-bottom:8px;">Week Numbering</div>
          ${['ISO Weeks (W23, W24…)','Project Weeks (W1, W2…)'].map((o,i)=>`
          <label style="display:flex;align-items:center;gap:8px;padding:4px 0;cursor:pointer;">
            <input type="radio" name="week" ${i===0?'checked':''} style="accent-color:${PROJECT.color};"/> <span style="font-size:13px;color:#323130;">${o}</span>
          </label>`).join('')}
        </div>

        <div style="margin-bottom:16px;">
          <div style="font-size:12px;font-weight:600;color:#323130;margin-bottom:8px;">Bar Style</div>
          ${['Gradient','Flat'].map((o,i)=>`
          <label style="display:flex;align-items:center;gap:8px;padding:4px 0;cursor:pointer;">
            <input type="radio" name="bar" ${i===0?'checked':''} style="accent-color:${PROJECT.color};"/> <span style="font-size:13px;color:#323130;">${o}</span>
          </label>`).join('')}
        </div>

        <div style="margin-bottom:16px;">
          <div style="font-size:12px;font-weight:600;color:#323130;margin-bottom:8px;">Row Height</div>
          ${['Compact (32px)','Normal (40px)','Spacious (52px)'].map((o,i)=>`
          <label style="display:flex;align-items:center;gap:8px;padding:4px 0;cursor:pointer;">
            <input type="radio" name="row" ${i===1?'checked':''} style="accent-color:${PROJECT.color};"/> <span style="font-size:13px;color:#323130;">${o}</span>
          </label>`).join('')}
        </div>

        <div>
          <div style="font-size:12px;font-weight:600;color:#323130;margin-bottom:8px;">Show / Hide</div>
          ${['Weekend shading','Dependency arrows','Progress % on bars','Assignee name on bars'].map((o,i)=>`
          <label style="display:flex;align-items:center;gap:8px;padding:5px 0;cursor:pointer;">
            <div style="width:32px;height:18px;background:${i<2?PROJECT.color:'#EDEBE9'};border-radius:9px;position:relative;flex-shrink:0;">
              <div style="width:14px;height:14px;background:white;border-radius:50%;position:absolute;top:2px;${i<2?'right:2px':'left:2px'};box-shadow:0 1px 3px rgba(0,0,0,0.2);"></div>
            </div>
            <span style="font-size:13px;color:#323130;">${o}</span>
          </label>`).join('')}
        </div>
      </div>
    </div>
  `);
}

// ── 6. Export menu ─────────────────────────────────────────────────────────
function exportMenuPage() {
  const svg = ganttSVG();
  return page(`
    ${toolbar('gantt')}
    <div style="position:relative;">
      <div style="overflow:auto;">${svg}</div>
      <!-- Callout positioned near top-right -->
      <div style="position:absolute;top:4px;right:12px;background:#fff;border:1px solid #EDEBE9;border-radius:4px;box-shadow:0 4px 16px rgba(0,0,0,0.12);min-width:210px;z-index:100;overflow:hidden;">
        <div style="padding:4px 0;">
          <div style="padding:9px 16px;font-size:13px;color:#323130;cursor:pointer;display:flex;align-items:center;gap:8px;">📥&ensp;Import Tasks…</div>
          <div style="padding:9px 16px;font-size:13px;color:#323130;cursor:pointer;display:flex;align-items:center;gap:8px;">📊&ensp;Export to Excel</div>
          <div style="padding:9px 16px;font-size:13px;color:#323130;cursor:pointer;background:#EFF6FC;display:flex;align-items:center;gap:8px;">📑&ensp;Export to PowerPoint</div>
          <div style="padding:9px 16px;font-size:13px;color:#323130;cursor:pointer;display:flex;align-items:center;gap:8px;">🖼&ensp;Export as Image (PNG)</div>
          <div style="height:1px;background:#EDEBE9;margin:4px 0;"></div>
          <div style="padding:9px 16px;font-size:13px;color:#323130;cursor:pointer;">✏️&ensp;Edit Project</div>
          <div style="padding:9px 16px;font-size:13px;color:#D13438;cursor:pointer;">🗑️&ensp;Delete Project</div>
        </div>
      </div>
    </div>
  `);
}

// ── Shared panel shell ─────────────────────────────────────────────────────
function panelShell(activeTab, titleText, content) {
  const svg = ganttSVG();
  const tab = t => `
    <button style="padding:7px 16px;font-size:13px;font-weight:${t===activeTab?600:400};
      color:${t===activeTab?PROJECT.color:'#605E5C'};
      border-bottom:${t===activeTab?`2px solid ${PROJECT.color}`:'2px solid transparent'};
      background:transparent;border:none;cursor:pointer;">${t.charAt(0).toUpperCase()+t.slice(1)}</button>`;
  return page(`
    ${toolbar('gantt')}
    <div style="display:flex;position:relative;overflow:hidden;">
      <div style="flex:1;overflow:auto;">${svg}</div>
      <div style="width:460px;flex-shrink:0;border-left:1px solid #EDEBE9;background:#fff;
                  display:flex;flex-direction:column;height:740px;">
        <!-- Panel header -->
        <div style="padding:16px 20px;border-bottom:1px solid #EDEBE9;display:flex;align-items:center;justify-content:space-between;">
          <span style="font-size:16px;font-weight:700;color:#323130;">${titleText}</span>
          <button style="border:none;background:none;font-size:18px;color:#605E5C;cursor:pointer;">✕</button>
        </div>
        <!-- Project banner -->
        <div style="display:flex;align-items:center;gap:8px;background:${PROJECT.color}12;
                    border-radius:4px;padding:8px 12px;margin:12px 20px 0;">
          <div style="width:10px;height:10px;border-radius:50%;background:${PROJECT.color};"></div>
          <span style="font-size:12px;color:#605E5C;">Project: <strong style="color:#323130;">${PROJECT.title}</strong></span>
        </div>
        <!-- Tabs -->
        <div style="display:flex;border-bottom:1px solid #EDEBE9;margin:10px 20px 0;padding:0;">
          ${tab('basic')}${tab('details')}${tab('links')}
        </div>
        <!-- Content -->
        <div style="flex:1;overflow:auto;padding:20px;">
          ${content}
        </div>
        <!-- Footer -->
        <div style="padding:14px 20px;border-top:1px solid #EDEBE9;display:flex;gap:8px;">
          <button class="btn btn-primary" style="min-width:120px;">Save Changes</button>
          <button class="btn btn-secondary">Cancel</button>
        </div>
      </div>
    </div>
  `);
}

// ── 7. Task panel — Details tab ────────────────────────────────────────────
function taskPanelDetailsPage() {
  const COLORS = ['#0078D4','#107C10','#CA5010','#8764B8','#038387','#D13438','#C43148','#00B7C3','#881798','#498205'];
  const selectedColor = '#8764B8'; // purple — shown as custom bar color for this task

  const swatches = [
    // "Auto" swatch
    `<div title="Auto color" style="width:28px;height:28px;border-radius:50%;
      background:linear-gradient(135deg,#ccc 50%,#fff 50%);cursor:pointer;
      border:2px solid #EDEBE9;box-sizing:border-box;flex-shrink:0;"></div>`,
    // Palette swatches
    ...COLORS.map(c => `
      <div style="width:28px;height:28px;border-radius:50%;background:${c};cursor:pointer;
        border:${c===selectedColor?'3px solid #323130':'2px solid transparent'};
        outline:${c===selectedColor?`2px solid ${c}`:'none'};
        outline-offset:2px;box-sizing:border-box;flex-shrink:0;"></div>`),
    // Rainbow custom swatch
    `<div title="Pick a custom color" style="width:28px;height:28px;border-radius:50%;
      background:conic-gradient(red,yellow,lime,cyan,blue,magenta,red);
      border:2px solid #EDEBE9;cursor:pointer;box-sizing:border-box;flex-shrink:0;"></div>`,
  ].join('');

  const content = `
    <div style="display:flex;flex-direction:column;gap:16px;">

      <!-- Phase -->
      <div>
        <label style="display:block;font-size:12px;font-weight:600;color:#323130;margin-bottom:4px;">Phase</label>
        <div style="position:relative;">
          <input value="Design" style="width:100%;padding:7px 10px;border:1px solid ${PROJECT.color};
            border-radius:4px;font-size:13px;font-family:inherit;color:#323130;box-sizing:border-box;"/>
        </div>
        <div style="font-size:11px;color:#605E5C;margin-top:4px;">
          Groups tasks visually on the Gantt. Start typing to see existing phases.
        </div>
      </div>

      <!-- Milestone toggle -->
      <div style="display:flex;align-items:center;justify-content:space-between;padding:4px 0;">
        <div>
          <div style="font-size:12px;font-weight:600;color:#323130;">Milestone</div>
          <div style="font-size:12px;color:#605E5C;margin-top:2px;">No</div>
        </div>
        <div style="width:44px;height:22px;background:#EDEBE9;border-radius:11px;position:relative;cursor:pointer;flex-shrink:0;">
          <div style="width:18px;height:18px;background:#fff;border-radius:50%;position:absolute;
            top:2px;left:2px;box-shadow:0 1px 3px rgba(0,0,0,0.2);"></div>
        </div>
      </div>

      <!-- Custom bar color -->
      <div>
        <label style="display:block;font-size:12px;font-weight:600;color:#323130;margin-bottom:6px;">
          Custom Bar Color <span style="font-weight:400;color:#605E5C;">(optional)</span>
        </label>
        <div style="display:flex;gap:7px;flex-wrap:wrap;align-items:center;">
          ${swatches}
        </div>
        <div style="display:flex;align-items:center;gap:6px;margin-top:6px;">
          <div style="width:10px;height:10px;border-radius:50%;background:${selectedColor};"></div>
          <span style="font-size:11px;color:#605E5C;">Custom color: ${selectedColor}</span>
        </div>
      </div>

      <!-- Notes -->
      <div>
        <label style="display:block;font-size:12px;font-weight:600;color:#323130;margin-bottom:4px;">Notes</label>
        <textarea style="width:100%;padding:7px 10px;border:1px solid #EDEBE9;border-radius:4px;
          font-size:13px;font-family:inherit;color:#323130;resize:none;height:108px;box-sizing:border-box;">Focus on mobile-first responsive layouts. Review with Sarah on brand guidelines before final approval. Figma source: figma.com/file/xyz</textarea>
      </div>

    </div>`;

  return panelShell('details', 'Edit: UX Wireframes', content);
}

// ── 8. Task panel — Links tab ──────────────────────────────────────────────
function taskPanelLinksPage() {
  // Show API Integration (task 8) which depends on UX Wireframes + Frontend Build
  const depChips = [
    { title: 'UX Wireframes',      status: 'In Progress', color: '#0078D4' },
    { title: 'Current State Analysis', status: 'In Progress', color: '#0078D4' },
  ].map(d => `
    <span style="display:inline-flex;align-items:center;gap:6px;padding:3px 8px 3px 10px;
      background:#EFF6FC;border:1px solid #90C8F6;border-radius:12px;font-size:12px;color:#0078D4;">
      <span style="width:6px;height:6px;border-radius:50%;background:${d.color};flex-shrink:0;"></span>
      ${d.title}
      <button style="background:none;border:none;cursor:pointer;padding:0 2px;
        color:#0078D4;font-size:14px;line-height:1;">×</button>
    </span>`).join('');

  const depOptions = ['Select a task…','Requirements Gathering','Visual Design','Frontend Build','QA Testing']
    .map((o,i) => `<option${i===0?' value=""':''}>${o}</option>`).join('');

  const parentOptions = [
    'None (top-level task)',
    'Requirements Gathering',
    'Current State Analysis',
    'UX Wireframes',
    'Frontend Build',
  ].map((o,i) => `<option${i===0?' selected':''}>${o}</option>`).join('');

  const content = `
    <div style="display:flex;flex-direction:column;gap:20px;">

      <!-- Parent Task -->
      <div>
        <label style="display:block;font-size:12px;font-weight:600;color:#323130;margin-bottom:4px;">Parent Task</label>
        <select style="width:100%;padding:7px 10px;border:1px solid #EDEBE9;border-radius:4px;
          font-size:13px;font-family:inherit;color:#323130;background:#fff;">
          ${parentOptions}
        </select>
        <div style="font-size:11px;color:#605E5C;margin-top:4px;">
          Makes this a sub-task, shown indented below the parent.
        </div>
      </div>

      <!-- Depends On -->
      <div>
        <label style="display:block;font-size:12px;font-weight:600;color:#323130;margin-bottom:6px;">Depends On</label>

        <!-- Chips -->
        <div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:10px;">
          ${depChips}
        </div>

        <!-- Add dependency dropdown -->
        <select style="width:100%;padding:7px 10px;border:1px solid #EDEBE9;border-radius:4px;
          font-size:13px;font-family:inherit;color:#605E5C;background:#fff;">
          ${depOptions}
        </select>

        <div style="font-size:11px;color:#605E5C;margin-top:6px;">
          This task cannot start until all dependencies are complete.
          Arrows are drawn on the Gantt chart.
        </div>
      </div>

    </div>`;

  return panelShell('links', 'Edit: API Integration', content);
}

// ── 9. PowerPoint export preview ──────────────────────────────────────────
// Renders a pixel-accurate HTML mock of the exported PPTX slides so the docs
// can show what the output looks like without needing LibreOffice to convert.
function pptxPreviewPage() {
  // Slide dimensions: 13.33" × 7.5" @ 96 dpi = 1280 × 720px
  const W = 1280, H = 720;
  const COLOR = PROJECT.color;
  const hex = c => c.replace('#','');

  // ── Stat data ───────────────────────────────────────────────────────────
  const byStatus = { 'Completed':2, 'In Progress':2, 'Not Started':5, 'On Hold':0, 'Cancelled':1 };
  const total = Object.values(byStatus).reduce((a,b)=>a+b,0);
  const overallPct = Math.round(TASKS.reduce((s,t)=>s+t.pct,0)/TASKS.length);

  const statItems = [
    { label:'Total Tasks',  value:total, bg:'#F3F2F1', fg:'#323130' },
    { label:'Completed',    value:byStatus['Completed'],   bg:'#F1FAF1', fg:'#107C10' },
    { label:'In Progress',  value:byStatus['In Progress'], bg:'#EFF6FC', fg:'#0078D4' },
    { label:'On Hold',      value:byStatus['On Hold'],     bg:'#FFF4EC', fg:'#CA5010' },
    { label:'Not Started',  value:byStatus['Not Started'], bg:'#F3F2F1', fg:'#605E5C' },
  ];

  const statusRows = [
    { label:'Completed',   color:'#107C10' },
    { label:'In Progress', color:'#0078D4' },
    { label:'Not Started', color:'#8B929A' },
    { label:'On Hold',     color:'#CA5010' },
    { label:'Cancelled',   color:'#D13438' },
  ];

  // ── Slide 1: Cover ──────────────────────────────────────────────────────
  const coverSlide = `
    <div style="width:${W}px;height:${H}px;background:${COLOR};position:relative;
                font-family:'Segoe UI',Arial,sans-serif;overflow:hidden;flex-shrink:0;">
      <!-- White lower panel -->
      <div style="position:absolute;bottom:0;left:0;right:0;height:220px;background:#fff;"></div>

      <!-- "PROJECT REPORT" label -->
      <div style="position:absolute;top:86px;left:62px;font-size:13px;color:rgba(255,255,255,0.7);
                  letter-spacing:3px;font-weight:400;">PROJECT REPORT</div>

      <!-- Project title -->
      <div style="position:absolute;top:110px;left:62px;right:62px;font-size:50px;font-weight:700;
                  color:#fff;line-height:1.15;">${PROJECT.title}</div>

      <!-- Status + dates -->
      <div style="position:absolute;top:290px;left:62px;font-size:18px;color:rgba(255,255,255,0.88);">
        Status: ${PROJECT.status}
      </div>
      <div style="position:absolute;top:320px;left:62px;font-size:15px;color:rgba(255,255,255,0.65);">
        Jun 1, 2026  →  Aug 31, 2026
      </div>

      <!-- Description (in white panel) -->
      <div style="position:absolute;bottom:140px;left:62px;right:62px;font-size:14px;color:#323130;line-height:1.5;">
        A full redesign of the company website including UX research, visual design,<br>
        frontend development, and QA — targeting a Q3 2026 launch.
      </div>

      <!-- Project Manager -->
      <div style="position:absolute;bottom:90px;left:62px;font-size:13px;color:#605E5C;">
        Project Manager: Sarah Miller
      </div>

      <!-- Generated date -->
      <div style="position:absolute;bottom:24px;right:24px;font-size:11px;color:#605E5C;">
        Generated June 4, 2026
      </div>
    </div>`;

  // ── Slide 2: Project Summary ────────────────────────────────────────────
  const barW = W - 100; // progress bar total width in px
  const fillW = Math.round(barW * overallPct / 100);

  const summarySlide = `
    <div style="width:${W}px;height:${H}px;background:#fff;position:relative;
                font-family:'Segoe UI',Arial,sans-serif;overflow:hidden;flex-shrink:0;">
      <!-- Header bar -->
      <div style="position:absolute;top:0;left:0;right:0;height:82px;background:${COLOR};
                  display:flex;align-items:center;padding:0 38px;justify-content:space-between;">
        <span style="font-size:24px;font-weight:700;color:#fff;">Project Summary</span>
        <span style="font-size:14px;color:rgba(255,255,255,0.65);">${PROJECT.title}</span>
      </div>

      <!-- Stat boxes -->
      <div style="position:absolute;top:105px;left:0;right:0;display:flex;justify-content:center;gap:14px;padding:0 32px;">
        ${statItems.map(s=>`
          <div style="flex:1;background:${s.bg};border:1px solid #EDEBE9;border-radius:6px;
                      padding:14px 10px 10px;text-align:center;">
            <div style="font-size:38px;font-weight:700;color:${s.fg};line-height:1;">${s.value}</div>
            <div style="font-size:12px;color:#605E5C;margin-top:8px;">${s.label}</div>
          </div>`).join('')}
      </div>

      <!-- Progress bar -->
      <div style="position:absolute;top:310px;left:50px;right:50px;">
        <div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:8px;">
          <span style="font-size:14px;font-weight:700;color:#323130;">Overall Progress</span>
          <span style="font-size:14px;font-weight:700;color:${COLOR};">${overallPct}%</span>
        </div>
        <div style="height:18px;background:#EDEBE9;border-radius:4px;overflow:hidden;">
          <div style="width:${overallPct}%;height:100%;background:${COLOR};border-radius:4px;"></div>
        </div>
      </div>

      <!-- Status breakdown -->
      <div style="position:absolute;top:385px;left:50px;right:50px;">
        <div style="font-size:13px;font-weight:700;color:#323130;margin-bottom:12px;">Status Breakdown</div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px 40px;">
          ${statusRows.map(s=>{
            const cnt = byStatus[s.label]||0;
            const pct = Math.round(cnt/total*100);
            return `<div style="display:flex;align-items:center;gap:8px;font-size:13px;color:#323130;">
              <div style="width:10px;height:10px;border-radius:50%;background:${s.color};flex-shrink:0;"></div>
              ${s.label}: <strong>${cnt}</strong> <span style="color:#605E5C;">(${pct}%)</span>
            </div>`;
          }).join('')}
        </div>
      </div>
    </div>`;

  // ── Slide 3: Gantt Timeline ─────────────────────────────────────────────
  // Embed the full gantt SVG, scaled to fit the slide width
  const svg = ganttSVG();
  const ganttSlide = `
    <div style="width:${W}px;height:${H}px;background:#fff;position:relative;
                font-family:'Segoe UI',Arial,sans-serif;overflow:hidden;flex-shrink:0;">
      <!-- Header bar -->
      <div style="position:absolute;top:0;left:0;right:0;height:72px;background:${COLOR};
                  display:flex;align-items:center;padding:0 38px;justify-content:space-between;z-index:1;">
        <span style="font-size:24px;font-weight:700;color:#fff;">Gantt Timeline</span>
        <span style="font-size:14px;color:rgba(255,255,255,0.65);">${PROJECT.title}</span>
      </div>
      <!-- Gantt SVG scaled to fit below the header -->
      <div style="position:absolute;top:72px;left:0;right:0;bottom:0;overflow:hidden;">
        <div style="transform:scale(0.78);transform-origin:top left;width:128%;">
          ${svg}
        </div>
      </div>
    </div>`;

  // ── Page layout: 3 slides stacked vertically with slide labels ──────────
  const slideCSS = `
    body { background:#6B6B6B; margin:0; padding:0; }
    .deck { display:flex; flex-direction:column; align-items:center; gap:0; padding:32px 40px; }
    .slide-wrap { display:flex; flex-direction:column; align-items:flex-start; margin-bottom:28px; }
    .slide-label { font-family:'Segoe UI',Arial,sans-serif; font-size:12px; color:rgba(255,255,255,0.7);
                   margin-bottom:6px; letter-spacing:0.5px; }
    .slide { box-shadow: 0 4px 24px rgba(0,0,0,0.5); }
  `;

  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>${slideCSS}</style></head><body>
<div class="deck">
  <div class="slide-wrap">
    <div class="slide-label">SLIDE 1 — Cover</div>
    <div class="slide">${coverSlide}</div>
  </div>
  <div class="slide-wrap">
    <div class="slide-label">SLIDE 2 — Project Summary</div>
    <div class="slide">${summarySlide}</div>
  </div>
  <div class="slide-wrap">
    <div class="slide-label">SLIDE 3 — Gantt Timeline</div>
    <div class="slide">${ganttSlide}</div>
  </div>
</div>
</body></html>`;
}

// ── Screenshot runner ──────────────────────────────────────────────────────
async function main() {
  const browser = await puppeteer.launch({
    executablePath: CHROME,
    headless: 'new',
    args: ['--no-sandbox','--disable-setuid-sandbox','--font-render-hinting=none'],
  });

  const shots = [
    ['screenshot-gantt.png',              ganttPage,              { width:1440, height:780 }],
    ['screenshot-list.png',               listPage,               { width:1440, height:660 }],
    ['screenshot-kanban.png',             kanbanPage,             { width:1440, height:800 }],
    ['screenshot-task-panel.png',         taskPanelPage,          { width:1440, height:820 }],
    ['screenshot-task-panel-details.png', taskPanelDetailsPage,   { width:1440, height:820 }],
    ['screenshot-task-panel-links.png',   taskPanelLinksPage,     { width:1440, height:820 }],
    ['screenshot-display-settings.png',   displaySettingsPage,    { width:1440, height:820 }],
    ['screenshot-export-menu.png',        exportMenuPage,         { width:1440, height:560 }],
    ['screenshot-pptx-export.png',        pptxPreviewPage,        { width:1360, height:2480 }],
  ];

  for (const [filename, htmlFn, vp] of shots) {
    const pg = await browser.newPage();
    await pg.setViewport(vp);
    await pg.setContent(htmlFn(), { waitUntil: 'load' });
    await pg.screenshot({ path: require('path').join(OUT, filename) });
    await pg.close();
    console.log('✓', filename);
  }

  await browser.close();
  console.log('\nAll screenshots saved to docs/screenshots/');
}

main().catch(e => { console.error(e); process.exit(1); });
