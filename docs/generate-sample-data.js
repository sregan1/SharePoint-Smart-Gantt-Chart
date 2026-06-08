'use strict';
// Generates a sample Excel file for importing into Smart Gantt Chart.
// Run with: node docs/generate-sample-data.js
// Output:   docs/sample-data/Tech-Conference-Tasks.xlsx

const XLSX = require('xlsx');
const path = require('path');
const fs   = require('fs');

// Column headers — every name matches an ImportService alias exactly.
// Auto-mapped on import with no manual column mapping required.
const COLUMNS = [
  'Task Name',    // → title
  'Phase',        // → phase
  'Start Date',   // → startDate
  'Due Date',     // → dueDate
  'Status',       // → status
  'Priority',     // → priority
  'Assigned To',  // → assignedTo
  'Email',        // → assignedToEmail
  '% Complete',   // → percentComplete
  'Description',  // → description
  'Notes',        // → notes
  'Milestone',    // → isMilestone  (Yes / blank)
];

// ── Task data ──────────────────────────────────────────────────────────────
// Scenario: "Annual Tech Conference Planning" — 25 tasks, 4 phases,
// dates spanning May–Oct 2026. Covers all importable fields and exercises
// every normalization path in ImportService (status, priority, %, milestone).
// Two rows intentionally leave Priority blank to test the "defaults to Medium" path.
const tasks = [
  // ── Planning (5 tasks) ──────────────────────────────────────────────────
  {
    'Task Name':   'Project Kickoff Meeting',
    'Phase':       'Planning',
    'Start Date':  '2026-05-07',
    'Due Date':    '2026-05-07',
    'Status':      'Completed',
    'Priority':    'High',
    'Assigned To': 'Sarah Chen',
    'Email':       's.chen@contoso.com',
    '% Complete':  100,
    'Description': 'Align all stakeholders on project goals, budget, and delivery timeline.',
    'Notes':       'All department leads attended; budget of $180k approved.',
    'Milestone':   'Yes',
  },
  {
    'Task Name':   'Define Conference Theme & Goals',
    'Phase':       'Planning',
    'Start Date':  '2026-05-08',
    'Due Date':    '2026-05-18',
    'Status':      'Completed',
    'Priority':    'High',
    'Assigned To': 'Sarah Chen',
    'Email':       's.chen@contoso.com',
    '% Complete':  100,
    'Description': 'Establish conference vision, key messages, and success metrics.',
    'Notes':       'Theme selected: "Innovate Together". Target: 500+ attendees.',
    'Milestone':   '',
  },
  {
    'Task Name':   'Venue Research & Shortlist',
    'Phase':       'Planning',
    'Start Date':  '2026-05-14',
    'Due Date':    '2026-05-25',
    'Status':      'Completed',
    'Priority':    'Critical',
    'Assigned To': 'Marcus Webb',
    'Email':       'm.webb@contoso.com',
    '% Complete':  100,
    'Description': 'Research venues with capacity 500+, AV capability, and in-house catering.',
    'Notes':       '3 venues shortlisted; Grand Convention Center ranked #1.',
    'Milestone':   '',
  },
  {
    'Task Name':   'Venue Selection & Contract Signed',
    'Phase':       'Planning',
    'Start Date':  '2026-05-28',
    'Due Date':    '2026-06-08',
    'Status':      'Completed',
    'Priority':    'Critical',
    'Assigned To': 'Marcus Webb',
    'Email':       'm.webb@contoso.com',
    '% Complete':  100,
    'Description': 'Finalize venue selection and execute contract.',
    'Notes':       'Grand Convention Center booked for Oct 1–2. Deposit paid.',
    'Milestone':   'Yes',
  },
  {
    'Task Name':   'Budget Approved by Finance',
    'Phase':       'Planning',
    'Start Date':  '2026-06-04',
    'Due Date':    '2026-06-15',
    'Status':      'Completed',
    'Priority':    'Critical',
    'Assigned To': 'Sarah Chen',
    'Email':       's.chen@contoso.com',
    '% Complete':  100,
    'Description': 'Secure full $180,000 project budget approval from CFO.',
    'Notes':       'Approved with 5% contingency reserve. Tracked in Finance portal.',
    'Milestone':   'Yes',
  },

  // ── Content (4 tasks) ───────────────────────────────────────────────────
  {
    'Task Name':   'Keynote Speaker Outreach',
    'Phase':       'Content',
    'Start Date':  '2026-06-11',
    'Due Date':    '2026-07-06',
    'Status':      'Completed',
    'Priority':    'High',
    'Assigned To': 'Priya Nair',
    'Email':       'p.nair@contoso.com',
    '% Complete':  100,
    'Description': 'Contact and confirm 3 keynote speakers for main stage.',
    'Notes':       'All 3 keynotes confirmed. Bios received and approved.',
    'Milestone':   '',
  },
  {
    'Task Name':   'Breakout Session Speaker Recruitment',
    'Phase':       'Content',
    'Start Date':  '2026-06-25',
    'Due Date':    '2026-07-20',
    'Status':      'Completed',
    'Priority':    'High',
    'Assigned To': 'Priya Nair',
    'Email':       'p.nair@contoso.com',
    '% Complete':  100,
    'Description': 'Recruit 12 speakers for 6 breakout sessions (2 per session).',
    'Notes':       '10 confirmed, 2 alternates confirmed as backup.',
    'Milestone':   '',
  },
  {
    'Task Name':   'Full Agenda Finalized & Published',
    'Phase':       'Content',
    'Start Date':  '2026-07-30',
    'Due Date':    '2026-08-10',
    'Status':      'Completed',
    'Priority':    'High',
    'Assigned To': 'Priya Nair',
    'Email':       'p.nair@contoso.com',
    '% Complete':  100,
    'Description': 'Publish complete agenda with speaker bios and session descriptions.',
    'Notes':       'Agenda live on website. PDF version distributed to sponsors.',
    'Milestone':   'Yes',
  },
  {
    'Task Name':   'Post-Event Survey Design',
    'Phase':       'Content',
    'Start Date':  '2026-08-27',
    'Due Date':    '2026-09-07',
    'Status':      'In Progress',
    'Priority':    'Medium',
    'Assigned To': 'Priya Nair',
    'Email':       'p.nair@contoso.com',
    '% Complete':  60,
    'Description': 'Design attendee satisfaction survey targeting NPS > 60.',
    'Notes':       'Draft in SurveyMonkey. Review with Sarah by Sep 4.',
    'Milestone':   '',
  },

  // ── Logistics (9 tasks) ─────────────────────────────────────────────────
  {
    'Task Name':   'AV & Production Vendor Selection',
    'Phase':       'Logistics',
    'Start Date':  '2026-06-18',
    'Due Date':    '2026-07-13',
    'Status':      'Completed',
    'Priority':    'High',
    'Assigned To': 'Marcus Webb',
    'Email':       'm.webb@contoso.com',
    '% Complete':  100,
    'Description': 'Select AV production company for main stage, 3 breakout rooms, and streaming.',
    'Notes':       'ProSound AV contracted. Includes live-stream package.',
    'Milestone':   '',
  },
  {
    'Task Name':   'Catering Menu & Contract',
    'Phase':       'Logistics',
    'Start Date':  '2026-06-25',
    'Due Date':    '2026-07-20',
    'Status':      'Completed',
    'Priority':    'Medium',
    'Assigned To': 'Marcus Webb',
    'Email':       'm.webb@contoso.com',
    '% Complete':  100,
    'Description': 'Finalize catering for 500 attendees across 2 full days.',
    'Notes':       'Menu confirmed. Halal, vegan, and gluten-free options included.',
    'Milestone':   '',
  },
  {
    'Task Name':   'Speaker Travel & Hotel Coordination',
    'Phase':       'Logistics',
    'Start Date':  '2026-08-06',
    'Due Date':    '2026-08-31',
    'Status':      'In Progress',
    'Priority':    'Medium',
    'Assigned To': 'Marcus Webb',
    'Email':       'm.webb@contoso.com',
    '% Complete':  60,
    'Description': 'Book flights and hotel accommodation for all out-of-town speakers.',
    'Notes':       '8 of 10 speakers fully booked. 2 awaiting flight confirmation.',
    'Milestone':   '',
  },
  {
    'Task Name':   'Badge & Swag Production',
    'Phase':       'Logistics',
    'Start Date':  '2026-08-13',
    'Due Date':    '2026-09-07',
    'Status':      'In Progress',
    'Priority':    'Low',
    'Assigned To': 'Marcus Webb',
    'Email':       'm.webb@contoso.com',
    '% Complete':  40,
    'Description': 'Design and print name badges; procure branded tote bags and lanyards.',
    'Notes':       'Badge design approved. Print order placed — delivery Sep 3.',
    'Milestone':   '',
  },
  {
    'Task Name':   'Volunteer Recruitment & Training',
    'Phase':       'Logistics',
    'Start Date':  '2026-08-20',
    'Due Date':    '2026-09-14',
    'Status':      'In Progress',
    'Priority':    'Medium',
    'Assigned To': 'Jess Okafor',
    'Email':       'j.okafor@contoso.com',
    '% Complete':  50,
    'Description': 'Recruit 30 event-day volunteers and run a 2-hour orientation session.',
    'Notes':       '22 volunteers confirmed. Orientation scheduled for Sep 11.',
    'Milestone':   '',
  },
  {
    'Task Name':   'Rehearsal & AV Tech Run-Through',
    'Phase':       'Logistics',
    'Start Date':  '2026-09-27',
    'Due Date':    '2026-09-28',
    'Status':      'Not Started',
    'Priority':    'Critical',
    'Assigned To': 'Marcus Webb',
    'Email':       'm.webb@contoso.com',
    '% Complete':  0,
    'Description': 'Full stage rehearsal with AV team and all keynote speakers.',
    'Notes':       'All speakers must attend. Backup AV team on standby.',
    'Milestone':   'Yes',
  },
  {
    'Task Name':   'Conference Day 1 Execution',
    'Phase':       'Logistics',
    'Start Date':  '2026-10-01',
    'Due Date':    '2026-10-01',
    'Status':      'Not Started',
    'Priority':    'Critical',
    'Assigned To': 'Sarah Chen',
    'Email':       's.chen@contoso.com',
    '% Complete':  0,
    'Description': 'Execute Day 1: registration open, 2 keynotes, 6 breakouts, evening reception.',
    'Notes':       'Doors open 7:30 AM. Live-stream goes live at 9 AM.',
    'Milestone':   'Yes',
  },
  {
    'Task Name':   'Conference Day 2 Execution',
    'Phase':       'Logistics',
    'Start Date':  '2026-10-02',
    'Due Date':    '2026-10-02',
    'Status':      'Not Started',
    'Priority':    'Critical',
    'Assigned To': 'Sarah Chen',
    'Email':       's.chen@contoso.com',
    '% Complete':  0,
    'Description': 'Execute Day 2: workshops, panel discussions, closing keynote, awards ceremony.',
    'Notes':       'Post-event survey email to go live at 4 PM.',
    'Milestone':   'Yes',
  },
  {
    'Task Name':   'Post-Event Venue Debrief',
    'Phase':       'Logistics',
    'Start Date':  '2026-10-08',
    'Due Date':    '2026-10-12',
    'Status':      'Not Started',
    // Priority intentionally blank — tests "defaults to Medium" in batchImport
    'Priority':    '',
    'Assigned To': 'Marcus Webb',
    'Email':       'm.webb@contoso.com',
    '% Complete':  0,
    'Description': 'Review venue performance with event manager; document lessons learned.',
    'Notes':       'Required for 2027 venue renewal negotiation.',
    'Milestone':   '',
  },

  // ── Marketing (7 tasks) ─────────────────────────────────────────────────
  {
    'Task Name':   'Registration Platform Setup',
    'Phase':       'Marketing',
    'Start Date':  '2026-07-09',
    'Due Date':    '2026-07-27',
    'Status':      'Completed',
    'Priority':    'High',
    'Assigned To': 'Jess Okafor',
    'Email':       'j.okafor@contoso.com',
    '% Complete':  100,
    'Description': 'Configure Eventbrite with early-bird and standard pricing tiers.',
    'Notes':       '312 early-bird registrations in first week.',
    'Milestone':   '',
  },
  {
    'Task Name':   'Conference Website Launch',
    'Phase':       'Marketing',
    'Start Date':  '2026-07-16',
    'Due Date':    '2026-08-03',
    'Status':      'Completed',
    'Priority':    'High',
    'Assigned To': 'Jess Okafor',
    'Email':       'j.okafor@contoso.com',
    '% Complete':  100,
    'Description': 'Launch public conference website with agenda, speakers, and sponsor info.',
    'Notes':       'Live at techconf2026.contoso.com. 4,200 unique visitors in week 1.',
    'Milestone':   'Yes',
  },
  {
    'Task Name':   'Social Media Campaign — Phase 1',
    'Phase':       'Marketing',
    'Start Date':  '2026-07-23',
    'Due Date':    '2026-08-17',
    'Status':      'Completed',
    'Priority':    'Medium',
    'Assigned To': 'Jess Okafor',
    'Email':       'j.okafor@contoso.com',
    '% Complete':  100,
    'Description': 'LinkedIn and Twitter campaign to drive early-bird ticket sales.',
    'Notes':       'Engagement rate 4.2% above baseline. 180 tickets sold via social.',
    'Milestone':   '',
  },
  {
    'Task Name':   'Sponsorship Packages Closed',
    'Phase':       'Marketing',
    'Start Date':  '2026-07-02',
    'Due Date':    '2026-08-24',
    'Status':      'In Progress',
    'Priority':    'High',
    'Assigned To': 'Sarah Chen',
    'Email':       's.chen@contoso.com',
    '% Complete':  80,
    'Description': 'Close Gold, Silver, and Bronze sponsorship packages.',
    'Notes':       '4 of 6 slots filled. 2 negotiations active — close expected by Aug 17.',
    'Milestone':   '',
  },
  {
    'Task Name':   'Social Media Campaign — Phase 2',
    'Phase':       'Marketing',
    'Start Date':  '2026-09-03',
    'Due Date':    '2026-09-30',
    'Status':      'Not Started',
    'Priority':    'Medium',
    'Assigned To': 'Jess Okafor',
    'Email':       'j.okafor@contoso.com',
    '% Complete':  0,
    'Description': 'Final countdown posts, speaker spotlights, and live-day social coverage.',
    'Notes':       'Coordinate speaker takeover content. Schedule posts via Buffer.',
    'Milestone':   '',
  },
  {
    'Task Name':   'Final Attendee Confirmation Emails',
    'Phase':       'Marketing',
    'Start Date':  '2026-09-17',
    'Due Date':    '2026-09-21',
    'Status':      'Not Started',
    // Priority intentionally blank — tests "defaults to Medium" in batchImport
    'Priority':    '',
    'Assigned To': 'Jess Okafor',
    'Email':       'j.okafor@contoso.com',
    '% Complete':  0,
    'Description': 'Send final logistics email to registered attendees with schedule and parking.',
    'Notes':       'Registration closes Sep 14. Use Mailchimp template #Conference-Final.',
    'Milestone':   '',
  },
  {
    'Task Name':   'Sponsor Debrief & 2027 Renewals',
    'Phase':       'Marketing',
    'Start Date':  '2026-10-03',
    'Due Date':    '2026-10-19',
    'Status':      'Not Started',
    'Priority':    'High',
    'Assigned To': 'Sarah Chen',
    'Email':       's.chen@contoso.com',
    '% Complete':  0,
    'Description': 'Present ROI metrics to all sponsors; pitch 2027 renewal packages.',
    'Notes':       'Pre-meeting decks due Oct 5. Target 80% renewal rate.',
    'Milestone':   '',
  },
];

// ── Write Excel file ───────────────────────────────────────────────────────
const ws = XLSX.utils.json_to_sheet(tasks, { header: COLUMNS });

// Auto-size columns
ws['!cols'] = COLUMNS.map(col => ({
  wch: Math.max(
    col.length + 2,
    ...tasks.map(row => String(row[col] ?? '').length + 1)
  ),
}));

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Tasks');

const outDir = path.join(__dirname, 'sample-data');
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

const outPath = path.join(outDir, 'Tech-Conference-Tasks.xlsx');
XLSX.writeFile(wb, outPath);

console.log(`Generated ${outPath}`);
console.log(`  ${tasks.length} tasks across ${[...new Set(tasks.map(t => t['Phase']))].length} phases`);
console.log(`  Statuses: ${[...new Set(tasks.map(t => t['Status']))].join(', ')}`);
console.log(`  Milestones: ${tasks.filter(t => t['Milestone'] === 'Yes').length}`);
console.log(`  Blank Priority rows (tests default): ${tasks.filter(t => t['Priority'] === '').length}`);
