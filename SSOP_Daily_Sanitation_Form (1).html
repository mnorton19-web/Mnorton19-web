<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SSOP — Daily Standard Sanitation | Royal Ridge Fruits</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root {
  --navy: #1a2e4a;
  --blue: #2563a8;
  --blue-light: #e8f0fb;
  --blue-mid: #b8d0f0;
  --green: #1a7a4a;
  --green-light: #e6f5ed;
  --amber: #92600a;
  --amber-light: #fef3db;
  --red: #9b2020;
  --red-light: #fdeaea;
  --gray-50: #f7f8fa;
  --gray-100: #eef0f3;
  --gray-200: #dde1e8;
  --gray-400: #9aa3b0;
  --gray-600: #5a6373;
  --gray-800: #2d3340;
  --white: #ffffff;
  --radius-sm: 6px;
  --radius-md: 10px;
  --radius-lg: 14px;
  --shadow-sm: 0 1px 3px rgba(0,0,0,0.07), 0 1px 2px rgba(0,0,0,0.05);
  --shadow-md: 0 4px 12px rgba(0,0,0,0.08), 0 2px 4px rgba(0,0,0,0.05);
}

* { box-sizing: border-box; margin: 0; padding: 0; }

body {
  font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
  font-size: 14px;
  background: #f0f2f5;
  color: var(--gray-800);
  min-height: 100vh;
}

/* ── Page nav ── */
.page { display: none; }
.page.active { display: block; }

/* ── FORM PAGE ── */
.form-page { max-width: 780px; margin: 0 auto; padding: 24px 16px 60px; }

.top-bar {
  background: var(--navy);
  color: white;
  padding: 14px 20px;
  border-radius: var(--radius-lg);
  margin-bottom: 20px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  flex-wrap: wrap;
  gap: 10px;
}
.top-bar-left h1 { font-size: 17px; font-weight: 600; letter-spacing: -0.2px; }
.top-bar-left p { font-size: 12px; opacity: 0.65; margin-top: 2px; }
.top-bar-right { display: flex; gap: 8px; flex-wrap: wrap; }

.btn {
  display: inline-flex; align-items: center; gap: 6px;
  padding: 8px 14px; border-radius: var(--radius-sm);
  font-size: 13px; font-weight: 500; cursor: pointer;
  border: none; transition: all 0.15s;
}
.btn-white { background: white; color: var(--navy); }
.btn-white:hover { background: #e8ecf2; }
.btn-outline { background: transparent; color: white; border: 1px solid rgba(255,255,255,0.35); }
.btn-outline:hover { background: rgba(255,255,255,0.1); }
.btn-green { background: var(--green); color: white; }
.btn-green:hover { background: #155e39; }
.btn-blue { background: var(--blue); color: white; }
.btn-blue:hover { background: #1d4f8a; }

/* Progress */
.progress-card {
  background: white;
  border-radius: var(--radius-md);
  padding: 14px 18px;
  margin-bottom: 16px;
  box-shadow: var(--shadow-sm);
  border: 1px solid var(--gray-200);
}
.progress-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
.progress-header span { font-size: 13px; color: var(--gray-600); }
.progress-header strong { font-size: 15px; color: var(--navy); }
.progress-track { height: 8px; background: var(--gray-100); border-radius: 4px; overflow: hidden; }
.progress-fill { height: 100%; background: linear-gradient(90deg, var(--blue), var(--green)); border-radius: 4px; transition: width 0.4s ease; }
.section-dots { display: flex; gap: 6px; margin-top: 10px; flex-wrap: wrap; }
.dot { width: 10px; height: 10px; border-radius: 50%; background: var(--gray-200); transition: background 0.3s; }
.dot.done { background: var(--green); }
.dot.partial { background: var(--blue); }

/* Meta card */
.meta-card {
  background: white;
  border-radius: var(--radius-md);
  border: 1px solid var(--gray-200);
  padding: 16px 18px;
  margin-bottom: 16px;
  box-shadow: var(--shadow-sm);
}
.meta-card h3 { font-size: 13px; font-weight: 600; color: var(--navy); margin-bottom: 12px; text-transform: uppercase; letter-spacing: 0.5px; }
.meta-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
.field label { display: block; font-size: 11px; font-weight: 600; color: var(--gray-600); text-transform: uppercase; letter-spacing: 0.4px; margin-bottom: 4px; }
.field input, .field select, .field textarea {
  width: 100%; padding: 8px 10px;
  border: 1px solid var(--gray-200); border-radius: var(--radius-sm);
  font-size: 13px; color: var(--gray-800); background: var(--gray-50);
  transition: border-color 0.15s;
  font-family: inherit;
}
.field input:focus, .field select:focus, .field textarea:focus {
  outline: none; border-color: var(--blue); background: white;
  box-shadow: 0 0 0 3px rgba(37,99,168,0.1);
}

/* Section cards */
.section-card {
  background: white;
  border-radius: var(--radius-md);
  border: 1px solid var(--gray-200);
  margin-bottom: 14px;
  box-shadow: var(--shadow-sm);
  overflow: hidden;
  transition: border-color 0.3s;
}
.section-card.complete { border-color: #a7d9ba; }

.section-head {
  display: flex; align-items: center; gap: 12px;
  padding: 13px 18px;
  background: var(--gray-50);
  border-bottom: 1px solid var(--gray-200);
}
.section-letter {
  width: 30px; height: 30px; border-radius: 50%;
  background: var(--blue); color: white;
  font-size: 13px; font-weight: 700;
  display: flex; align-items: center; justify-content: center;
  flex-shrink: 0; transition: background 0.3s;
}
.section-card.complete .section-letter { background: var(--green); }
.section-info { flex: 1; min-width: 0; }
.section-info h3 { font-size: 14px; font-weight: 600; color: var(--navy); }
.section-info p { font-size: 11px; color: var(--gray-400); margin-top: 1px; }
.section-count {
  font-size: 12px; font-weight: 600; padding: 3px 10px;
  border-radius: 20px; background: var(--gray-100); color: var(--gray-600);
  white-space: nowrap;
}
.section-card.complete .section-count { background: var(--green-light); color: var(--green); }

.section-body { padding: 0 18px 14px; }

/* Technician field per section */
.tech-row {
  display: flex; align-items: center; gap: 10px;
  padding: 10px 0 10px;
  border-bottom: 1px solid var(--gray-100);
  margin-bottom: 4px;
}
.tech-row label { font-size: 11px; font-weight: 600; color: var(--gray-600); text-transform: uppercase; letter-spacing: 0.4px; white-space: nowrap; }
.tech-row input { flex: 1; padding: 6px 10px; border: 1px solid var(--gray-200); border-radius: var(--radius-sm); font-size: 13px; background: var(--gray-50); color: var(--gray-800); font-family: inherit; }
.tech-row input:focus { outline: none; border-color: var(--blue); background: white; box-shadow: 0 0 0 3px rgba(37,99,168,0.1); }

/* Task rows */
.task-row {
  display: flex; align-items: flex-start; gap: 10px;
  padding: 8px 0; border-bottom: 1px solid var(--gray-100);
}
.task-row:last-of-type { border-bottom: none; }
.task-row input[type=checkbox] {
  width: 16px; height: 16px; margin-top: 2px; flex-shrink: 0;
  accent-color: var(--green); cursor: pointer;
}
.task-text { flex: 1; font-size: 13px; line-height: 1.5; cursor: pointer; }
.task-text.done { text-decoration: line-through; color: var(--gray-400); }
.task-ref { font-size: 10px; color: var(--gray-400); align-self: center; white-space: nowrap; flex-shrink: 0; }

/* Alert */
.alert {
  display: flex; gap: 8px; padding: 9px 12px;
  border-radius: var(--radius-sm); font-size: 12px; line-height: 1.5;
  margin: 10px 0 4px;
}
.alert-warn { background: var(--amber-light); color: var(--amber); border-left: 3px solid #e09a20; }
.alert-info { background: var(--blue-light); color: var(--blue); border-left: 3px solid var(--blue); }
.alert-success { background: var(--green-light); color: var(--green); border-left: 3px solid var(--green); }
.alert-danger { background: var(--red-light); color: var(--red); border-left: 3px solid var(--red); }
.alert svg { flex-shrink: 0; margin-top: 1px; }

.notes-row { margin-top: 8px; }
.notes-row label { font-size: 11px; font-weight: 600; color: var(--gray-600); text-transform: uppercase; letter-spacing: 0.4px; display: block; margin-bottom: 4px; }
.notes-row textarea { width: 100%; padding: 8px 10px; border: 1px solid var(--gray-200); border-radius: var(--radius-sm); font-size: 12px; background: var(--gray-50); color: var(--gray-800); resize: vertical; min-height: 52px; font-family: inherit; }
.notes-row textarea:focus { outline: none; border-color: var(--blue); background: white; box-shadow: 0 0 0 3px rgba(37,99,168,0.1); }

/* Sign-off */
.signoff-card {
  background: white; border-radius: var(--radius-md);
  border: 1px solid var(--gray-200); padding: 16px 18px;
  margin-bottom: 16px; box-shadow: var(--shadow-sm);
}
.signoff-card h3 { font-size: 13px; font-weight: 600; color: var(--navy); margin-bottom: 12px; text-transform: uppercase; letter-spacing: 0.5px; }

/* Submit bar */
.submit-bar {
  background: white; border-radius: var(--radius-md);
  border: 1px solid var(--gray-200); padding: 14px 18px;
  display: flex; gap: 10px; flex-wrap: wrap; align-items: center;
  box-shadow: var(--shadow-sm);
}
.submit-bar .spacer { flex: 1; }

/* Success banner */
.success-banner {
  display: none; background: var(--green-light);
  border: 1px solid #a7d9ba; border-radius: var(--radius-md);
  padding: 16px 18px; margin-top: 16px; text-align: center;
}
.success-banner h2 { font-size: 16px; font-weight: 600; color: var(--green); }
.success-banner p { font-size: 13px; color: var(--green); opacity: 0.85; margin-top: 4px; }

/* ── QR PAGE ── */
.qr-page {
  max-width: 500px; margin: 40px auto; padding: 16px;
  text-align: center;
}
.qr-card {
  background: white; border-radius: var(--radius-lg);
  border: 1px solid var(--gray-200); padding: 32px 24px;
  box-shadow: var(--shadow-md);
}
.qr-card .logo-bar {
  background: var(--navy); color: white;
  border-radius: var(--radius-md); padding: 10px 16px;
  margin-bottom: 24px; font-size: 13px; font-weight: 600;
}
#qrcode { display: inline-block; margin: 0 auto 20px; }
#qrcode canvas, #qrcode img { border-radius: 8px; }
.qr-url { font-size: 11px; color: var(--gray-400); word-break: break-all; margin-bottom: 20px; }
.qr-title { font-size: 18px; font-weight: 700; color: var(--navy); margin-bottom: 6px; }
.qr-sub { font-size: 13px; color: var(--gray-600); margin-bottom: 20px; }
.qr-steps {
  text-align: left; background: var(--gray-50);
  border-radius: var(--radius-md); padding: 14px 16px; margin-bottom: 20px;
}
.qr-steps p { font-size: 12px; color: var(--gray-600); line-height: 1.8; }
.qr-btn-row { display: flex; justify-content: center; gap: 10px; flex-wrap: wrap; }

/* ── RECORDS PAGE ── */
.records-page { max-width: 780px; margin: 0 auto; padding: 24px 16px 60px; }
.records-top { display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; flex-wrap: wrap; gap: 10px; }
.records-top h2 { font-size: 18px; font-weight: 600; color: var(--navy); }
.records-table-wrap { background: white; border-radius: var(--radius-md); border: 1px solid var(--gray-200); overflow: auto; box-shadow: var(--shadow-sm); }
table { width: 100%; border-collapse: collapse; font-size: 12px; }
th { background: var(--navy); color: white; padding: 10px 12px; text-align: left; font-weight: 600; font-size: 11px; white-space: nowrap; }
td { padding: 9px 12px; border-bottom: 1px solid var(--gray-100); vertical-align: top; }
tr:hover td { background: var(--gray-50); }
.badge {
  display: inline-block; padding: 2px 8px; border-radius: 20px;
  font-size: 11px; font-weight: 600;
}
.badge-complete { background: var(--green-light); color: var(--green); }
.badge-partial { background: var(--amber-light); color: var(--amber); }
.empty-state { text-align: center; padding: 40px; color: var(--gray-400); font-size: 13px; }

/* Nav tabs */
.nav-tabs { display: flex; gap: 4px; background: white; padding: 6px; border-radius: var(--radius-md); border: 1px solid var(--gray-200); margin-bottom: 20px; box-shadow: var(--shadow-sm); }
.nav-tab { flex: 1; padding: 8px; text-align: center; border-radius: var(--radius-sm); font-size: 13px; font-weight: 500; cursor: pointer; color: var(--gray-600); transition: all 0.15s; border: none; background: none; }
.nav-tab.active { background: var(--navy); color: white; }
.nav-tab:hover:not(.active) { background: var(--gray-100); }

/* Inline icons (simple SVG) */
.icon { display: inline-block; width: 14px; height: 14px; vertical-align: -2px; }

@media (max-width: 500px) {
  .meta-grid { grid-template-columns: 1fr; }
  .sig-grid { grid-template-columns: 1fr; }
  .top-bar { flex-direction: column; align-items: flex-start; }
}

@media print {
  .nav-tabs, .top-bar-right, .submit-bar, .success-banner { display: none !important; }
  body { background: white; }
  .form-page { padding: 0; max-width: 100%; }
  .section-card, .meta-card, .progress-card, .signoff-card { box-shadow: none; break-inside: avoid; }
}
</style>
</head>
<body>

<div style="max-width:780px;margin:0 auto;padding:16px 16px 0">
  <div class="nav-tabs">
    <button class="nav-tab active" onclick="showPage('form')">&#x1F4CB; Daily Form</button>
    <button class="nav-tab" onclick="showPage('records')">&#x1F4CA; Records</button>
    <button class="nav-tab" onclick="showPage('qr')">&#x1F4F1; QR Code</button>
  </div>
</div>

<!-- ══════════════════════════════════════════════
     FORM PAGE
══════════════════════════════════════════════ -->
<div class="page active" id="page-form">
<div class="form-page">

  <div class="top-bar">
    <div class="top-bar-left">
      <h1>&#x1F9FC; SSOP — Daily Standard Sanitation</h1>
      <p>Royal Ridge Fruits &nbsp;|&nbsp; SSOP-SAN-001 &nbsp;|&nbsp; SQF Ed. 9 / FSMA PC Compliant</p>
    </div>
    <div class="top-bar-right">
      <button class="btn btn-outline" onclick="window.print()">&#x1F5A8; Print</button>
      <button class="btn btn-white" onclick="resetForm()">&#x21BA; Reset</button>
    </div>
  </div>

  <!-- Progress -->
  <div class="progress-card">
    <div class="progress-header">
      <span>Overall completion</span>
      <strong id="prog-label">0 of 0 tasks</strong>
    </div>
    <div class="progress-track"><div class="progress-fill" id="prog-fill" style="width:0%"></div></div>
    <div class="section-dots" id="section-dots"></div>
  </div>

  <!-- Meta -->
  <div class="meta-card">
    <h3>Shift Details</h3>
    <div class="meta-grid">
      <div class="field"><label>Date</label><input type="date" id="f-date"></div>
      <div class="field"><label>Shift</label>
        <select id="f-shift"><option value="">Select...</option><option>Day</option><option>Swing</option><option>Night</option></select>
      </div>
      <div class="field"><label>Production line / area</label><input type="text" id="f-area" placeholder="e.g. Line 1, Packhouse A"></div>
      <div class="field"><label>Supervisor name</label><input type="text" id="f-sup" placeholder="Full name"></div>
    </div>
  </div>

  <!-- Sections injected here -->
  <div id="sections-wrap"></div>

  <!-- Sign-off -->
  <div class="signoff-card">
    <h3>Supervisor Sign-Off</h3>
    <div class="alert alert-info">
      <svg class="icon" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
      All sections must be completed and initialled before production starts. Records retained 2 years per 21 CFR 117.190.
    </div>
    <div class="meta-grid sig-grid" style="margin-top:12px">
      <div class="field"><label>Supervisor initials</label><input type="text" id="sig-initials" placeholder="e.g. MN" maxlength="8"></div>
      <div class="field"><label>Verification time</label><input type="time" id="sig-time"></div>
      <div class="field" style="grid-column:1/-1">
        <label>Corrective actions / deviations noted</label>
        <textarea id="sig-ca" rows="3" placeholder="Note any deviations, corrective actions taken, or items added to the CAL..."></textarea>
      </div>
    </div>
  </div>

  <div class="submit-bar">
    <div id="submit-status" style="font-size:13px;color:var(--gray-600)"></div>
    <div class="spacer"></div>
    <button class="btn btn-blue" onclick="exportExcel()">&#x1F4E5; Export to Excel</button>
    <button class="btn btn-green" onclick="submitForm()">&#x2713; Submit Record</button>
  </div>

  <div class="success-banner" id="success-banner">
    <h2>&#x2705; Sanitation record submitted</h2>
    <p id="success-msg"></p>
  </div>

</div>
</div>

<!-- ══════════════════════════════════════════════
     RECORDS PAGE
══════════════════════════════════════════════ -->
<div class="page" id="page-records">
<div class="records-page">
  <div class="records-top">
    <h2>&#x1F4CA; Submission Records</h2>
    <div style="display:flex;gap:8px">
      <button class="btn btn-blue" onclick="exportAllExcel()">&#x1F4E5; Export All to Excel</button>
      <button class="btn btn-outline" style="background:var(--red-light);color:var(--red);border:1px solid #e8aaaa" onclick="clearRecords()">&#x1F5D1; Clear All</button>
    </div>
  </div>
  <div class="records-table-wrap">
    <table id="records-table">
      <thead>
        <tr>
          <th>#</th><th>Date</th><th>Shift</th><th>Area</th><th>Supervisor</th>
          <th>Tasks</th><th>Status</th><th>Submitted</th>
        </tr>
      </thead>
      <tbody id="records-body"><tr><td colspan="8" class="empty-state">No records yet. Submit a completed form to see it here.</td></tr></tbody>
    </table>
  </div>
</div>
</div>

<!-- ══════════════════════════════════════════════
     QR PAGE
══════════════════════════════════════════════ -->
<div class="page" id="page-qr">
<div class="qr-page">
  <div class="qr-card">
    <div class="logo-bar">Royal Ridge Fruits &nbsp;|&nbsp; SSOP-SAN-001</div>
    <p class="qr-title">Daily Sanitation Form</p>
    <p class="qr-sub">Scan to open on any device</p>
    <div id="qrcode"></div>
    <p class="qr-url" id="qr-url-display"></p>
    <div class="qr-steps">
      <p><strong>How to use this QR code:</strong></p>
      <p>1. Host this HTML file on a shared drive, intranet, or web server.</p>
      <p>2. Come back to this QR tab — it auto-generates from the current URL.</p>
      <p>3. Print this page and post it in the sanitation staging area.</p>
      <p>4. Technicians scan with any phone to open the form.</p>
    </div>
    <div class="qr-btn-row">
      <button class="btn btn-blue" onclick="window.print()">&#x1F5A8; Print QR Page</button>
      <button class="btn" style="background:var(--gray-100);color:var(--gray-800)" onclick="downloadQR()">&#x1F4BE; Save QR Image</button>
    </div>
  </div>
</div>
</div>

<script>
// ── Section data ──────────────────────────────────────────────────────────
const SECTIONS = [
  {id:"A", title:"Trash removal & waste control", ref:"FSMA 117.135 / SQF 2.4.4",
   warn:"Do not allow trash to overflow. Overflow is a Preventive Control deviation under 21 CFR 117.150.",
   tasks:[
    ["Empty all production trash bins.","SQF 2.4.4"],
    ["Empty breakroom trash daily or more frequently as needed.","SQF 2.4.4"],
    ["Empty restroom waste containers daily or as needed.","SQF 2.4.4"],
    ["Transport all waste immediately to the designated exterior waste area.","SQF 2.4.4"],
    ["Ensure waste is contained in approved, labeled bins.","FSMA 117.135"],
    ["Replace bin liners immediately after each removal.","SQF 2.4.4"],
    ["Clean and sanitize trash bins before returning to production zones.","FSMA 117.135"],
    ["Verify waste handling does not create contamination risk to food contact areas.","FSMA 117.135"],
  ]},
  {id:"B", title:"Foot mats & hygiene barriers", ref:"SQF 2.4.4 / FSMA 117.145",
   warn:"Sanitizer concentrations outside validated range are a Preventive Control deviation — correct and re-document before production.",
   tasks:[
    ["Remove and scrub sanitation mats one at a time with approved cleaner.","SQF 2.4.4"],
    ["Rinse mats thoroughly; verify no chemical residue remains.","FSMA 117.135"],
    ["Return mats to proper location; verify correct seating (no trip hazard).","SQF 2.4.4"],
    ["Refill footbath mats with correct sanitizer at approved concentration; document concentration.","FSMA 117.145"],
    ["Apply P-Quat floor cleaner at every production area transition; reapply as needed.","FSMA 117.135"],
    ["Verify mat concentration is within validated range; log on Sanitation Monitoring Record.","FSMA 117.145"],
  ]},
  {id:"C", title:"Outside grounds control", ref:"SQF 2.4.4 / FSMA 117.135",
   warn:"Pest evidence must be documented on the Pest Activity Log and escalated to the Sanitation Manager immediately.",
   tasks:[
    ["Inspect and clean all entryways and door thresholds.","SQF 2.4.4"],
    ["Inspect and clean waste areas and dumpster pads.","SQF 2.4.4"],
    ["Inspect and clean pallet staging areas; remove debris.","SQF 2.4.4"],
    ["Inspect building perimeter (high-risk zones) for pest evidence.","SQF 2.4.4"],
    ["Remove all loose cardboard from exterior.","SQF 2.4.4"],
    ["Remove fruit waste, standing water, and standing debris.","FSMA 117.135"],
    ["Document pest evidence on Pest Activity Log; notify PCO if activity observed.","FSMA 117.150"],
  ]},
  {id:"D", title:"Janitorial — non-production areas", ref:"SQF 2.4.3 / SQF 2.4.4",
   warn:"Non-functional hand-washing stations must be corrected before production begins (SQF 2.4.3).",
   tasks:[
    ["Sweep and wet-mop breakroom floors with approved sanitizer.","SQF 2.4.4"],
    ["Wipe and sanitize all breakroom surfaces (tables, countertops, handles).","SQF 2.4.4"],
    ["Sweep and wet-mop locker room floors.","SQF 2.4.4"],
    ["Clean and sanitize restrooms — floors, fixtures, handles, and surfaces.","SQF 2.4.4"],
    ["Refill soap dispensers, paper towels, and toilet paper in all areas.","SQF 2.4.3"],
    ["Inspect and clean offices as required.","SQF 2.4.4"],
    ["Verify hand-washing stations are operational and stocked.","SQF 2.4.3 / FSMA 117.135"],
  ]},
  {id:"E", title:"Next-shift sanitation preparation", ref:"FSMA 117.135 / FSMA 117.145",
   warn:"No shift may begin without completed sanitation readiness verification. Failure is a Preventive Control deviation under 21 CFR 117.150.",
   tasks:[
    ["Inspect all hoses for cracks, leaks, or damage (harborage risk).","FSMA 117.135"],
    ["Stage hoses — properly coiled, off the floor, stored in designated location.","SQF 2.4.4"],
    ["Stage foaming units — filled, functional, calibrated to correct dilution.","FSMA 117.145"],
    ["Verify chemical inventory — levels sufficient for next full shift.","SQF 2.4.4"],
    ["Organize tools — clean, sanitized, and sorted by color-code designation.","SQF 2.4.4"],
    ["Confirm all equipment is functional; tag defective equipment Out of Service.","FSMA 117.135"],
    ["Confirm work areas are organized and inspection-ready.","SQF 2.6.1"],
    ["Supervisor initials Sanitation Readiness Record before close of shift.","FSMA 117.145"],
  ]},
  {id:"F", title:"Daily MSS task", ref:"FSMA 117.145 / SQF 2.6.1",
   warn:"MSS task must be logged with description, employee name, date, and supervisor initials. Retained 2 years per FSMA.",
   tasks:[
    ["Assign and complete one MSS task (overhead, drains, structural, equipment deep clean, hard-to-reach, or cold storage).","SQF 2.4.4"],
    ["Document MSS task on MSS Tracking Log — task, area, employee, date.","FSMA 117.145"],
    ["Supervisor verifies and initials MSS Tracking Log entry.","SQF 2.6.1"],
  ]},
];

// ── State ─────────────────────────────────────────────────────────────────
let checks = {};
let totalTasks = 0;
let records = JSON.parse(localStorage.getItem("ssop_records") || "[]");

// ── Build form sections ───────────────────────────────────────────────────
function buildSections() {
  const wrap = document.getElementById("sections-wrap");
  const dots = document.getElementById("section-dots");
  SECTIONS.forEach(sec => {
    totalTasks += sec.tasks.length;

    // dot
    const dot = document.createElement("span");
    dot.className = "dot"; dot.id = "dot-" + sec.id;
    dot.title = sec.title;
    dots.appendChild(dot);

    const card = document.createElement("div");
    card.className = "section-card"; card.id = "sec-" + sec.id;

    // header
    card.innerHTML = `
      <div class="section-head">
        <div class="section-letter" id="ltr-${sec.id}">${sec.id}</div>
        <div class="section-info">
          <h3>${sec.title}</h3>
          <p>${sec.ref}</p>
        </div>
        <span class="section-count" id="cnt-${sec.id}">0 / ${sec.tasks.length}</span>
      </div>`;

    const body = document.createElement("div");
    body.className = "section-body";

    // technician field
    body.innerHTML = `
      <div class="tech-row">
        <label>Technician</label>
        <input type="text" id="tech-${sec.id}" placeholder="Name of technician completing this section">
      </div>`;

    // tasks
    sec.tasks.forEach((t, i) => {
      const key = sec.id + "_" + i;
      checks[key] = false;
      const row = document.createElement("div");
      row.className = "task-row";
      row.innerHTML = `
        <input type="checkbox" id="chk-${key}" onchange="toggle('${key}','${sec.id}',${sec.tasks.length})">
        <label class="task-text" id="lbl-${key}" for="chk-${key}">${t[0]}</label>
        <span class="task-ref">${t[1]}</span>`;
      body.appendChild(row);
    });

    // warning
    const warn = document.createElement("div");
    warn.className = "alert alert-warn";
    warn.innerHTML = `<svg class="icon" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"/></svg><span>${sec.warn}</span>`;
    body.appendChild(warn);

    // notes
    const notesDiv = document.createElement("div");
    notesDiv.className = "notes-row";
    notesDiv.innerHTML = `<label>Section notes / deviations</label><textarea id="notes-${sec.id}" placeholder="Observations, deviations, or follow-up items for this section..."></textarea>`;
    body.appendChild(notesDiv);

    card.appendChild(body);
    wrap.appendChild(card);
  });
  updateProgress();
}

function toggle(key, secId, total) {
  checks[key] = document.getElementById("chk-" + key).checked;
  document.getElementById("lbl-" + key).className = checks[key] ? "task-text done" : "task-text";
  updateSection(secId, total);
  updateProgress();
}

function updateSection(secId, total) {
  const done = Object.keys(checks).filter(k => k.startsWith(secId + "_") && checks[k]).length;
  document.getElementById("cnt-" + secId).textContent = done + " / " + total;
  const complete = done === total;
  const card = document.getElementById("sec-" + secId);
  const dot = document.getElementById("dot-" + secId);
  card.className = complete ? "section-card complete" : "section-card";
  dot.className = complete ? "dot done" : (done > 0 ? "dot partial" : "dot");
}

function updateProgress() {
  const done = Object.values(checks).filter(Boolean).length;
  const pct = totalTasks > 0 ? Math.round((done / totalTasks) * 100) : 0;
  document.getElementById("prog-fill").style.width = pct + "%";
  document.getElementById("prog-label").textContent = done + " of " + totalTasks + " tasks (" + pct + "%)";
  const status = document.getElementById("submit-status");
  if (done === totalTasks) {
    status.innerHTML = '<span style="color:var(--green);font-weight:600">&#x2713; All tasks complete — ready to submit</span>';
  } else {
    status.textContent = (totalTasks - done) + " tasks remaining";
  }
}

// ── Submit ────────────────────────────────────────────────────────────────
function submitForm() {
  const date = document.getElementById("f-date").value;
  const shift = document.getElementById("f-shift").value;
  const sup = document.getElementById("f-sup").value;
  const initials = document.getElementById("sig-initials").value;
  if (!date || !shift || !sup || !initials) {
    alert("Please fill in Date, Shift, Supervisor name, and Supervisor initials before submitting.");
    return;
  }
  const done = Object.values(checks).filter(Boolean).length;
  if (done < totalTasks) {
    if (!confirm(done + " of " + totalTasks + " tasks checked. Submit with incomplete tasks? Unchecked items will be flagged as deviations.")) return;
  }

  const record = buildRecord();
  records.push(record);
  try { localStorage.setItem("ssop_records", JSON.stringify(records)); } catch(e) {}
  renderRecords();

  const banner = document.getElementById("success-banner");
  const d = new Date(date + "T12:00:00").toLocaleDateString("en-US", {year:"numeric",month:"long",day:"numeric"});
  document.getElementById("success-msg").textContent =
    shift + " shift — " + d + " — Supervisor: " + sup + " (" + initials + ") — " + done + "/" + totalTasks + " tasks completed.";
  banner.style.display = "block";
  banner.scrollIntoView({behavior:"smooth"});
}

function buildRecord() {
  const date = document.getElementById("f-date").value;
  const shift = document.getElementById("f-shift").value;
  const area = document.getElementById("f-area").value;
  const sup = document.getElementById("f-sup").value;
  const initials = document.getElementById("sig-initials").value;
  const sigTime = document.getElementById("sig-time").value;
  const ca = document.getElementById("sig-ca").value;
  const done = Object.values(checks).filter(Boolean).length;
  const sectionData = {};
  SECTIONS.forEach(sec => {
    const secDone = Object.keys(checks).filter(k => k.startsWith(sec.id + "_") && checks[k]).length;
    sectionData[sec.id] = {
      technician: document.getElementById("tech-" + sec.id).value,
      done: secDone,
      total: sec.tasks.length,
      notes: document.getElementById("notes-" + sec.id).value,
      tasks: sec.tasks.map((t, i) => ({ task: t[0], ref: t[1], done: checks[sec.id + "_" + i] || false }))
    };
  });
  return { date, shift, area, supervisor: sup, initials, sigTime, ca, tasksTotal: totalTasks, tasksDone: done, sections: sectionData, submitted: new Date().toISOString() };
}

// ── Export single record to Excel ─────────────────────────────────────────
function exportExcel() {
  const date = document.getElementById("f-date").value;
  if (!date) { alert("Please enter a date before exporting."); return; }
  const record = buildRecord();
  exportRecordToExcel([record], "SSOP_" + date);
}

function exportAllExcel() {
  if (records.length === 0) { alert("No records to export yet."); return; }
  exportRecordToExcel(records, "SSOP_All_Records");
}

function exportRecordToExcel(recs, filename) {
  const wb = XLSX.utils.book_new();

  // ── Sheet 1: Summary ──
  const summaryRows = [
    ["Royal Ridge Fruits — SSOP Daily Standard Sanitation"],
    ["SSOP-SAN-001 | SQF Ed. 9 / FSMA PC Compliant"],
    [],
    ["Date","Shift","Area","Supervisor","Initials","Time","Tasks Done","Tasks Total","% Complete","Status","Corrective Actions","Submitted"]
  ];
  recs.forEach(r => {
    const pct = Math.round((r.tasksDone / r.tasksTotal) * 100);
    summaryRows.push([
      r.date, r.shift, r.area || "", r.supervisor, r.initials, r.sigTime || "",
      r.tasksDone, r.tasksTotal, pct + "%",
      r.tasksDone === r.tasksTotal ? "Complete" : "Incomplete — Deviations Present",
      r.ca || "",
      new Date(r.submitted).toLocaleString()
    ]);
  });
  const ws1 = XLSX.utils.aoa_to_sheet(summaryRows);
  ws1["!cols"] = [14,8,16,18,8,8,10,10,10,22,30,20].map(w => ({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws1, "Summary");

  // ── Sheet 2: Task Detail ──
  const detailRows = [
    ["Royal Ridge Fruits — SSOP Task Detail"],
    [],
    ["Date","Shift","Supervisor","Section","Section Title","Technician","Task","Reference","Completed","Section Notes"]
  ];
  recs.forEach(r => {
    SECTIONS.forEach(sec => {
      const sd = r.sections[sec.id];
      if (!sd) return;
      sd.tasks.forEach(t => {
        detailRows.push([
          r.date, r.shift, r.supervisor,
          sec.id, sec.title, sd.technician || "",
          t.task, t.ref,
          t.done ? "Yes" : "No",
          sd.notes || ""
        ]);
      });
    });
  });
  const ws2 = XLSX.utils.aoa_to_sheet(detailRows);
  ws2["!cols"] = [12,8,18,6,28,18,55,22,10,30].map(w => ({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws2, "Task Detail");

  // ── Sheet 3: Section Summary ──
  const secRows = [
    ["Royal Ridge Fruits — Section Summary"],
    [],
    ["Date","Shift","Supervisor","Section","Section Title","Technician","Tasks Done","Tasks Total","Complete","Notes"]
  ];
  recs.forEach(r => {
    SECTIONS.forEach(sec => {
      const sd = r.sections[sec.id];
      if (!sd) return;
      secRows.push([
        r.date, r.shift, r.supervisor,
        sec.id, sec.title, sd.technician || "",
        sd.done, sd.total,
        sd.done === sd.total ? "Yes" : "No",
        sd.notes || ""
      ]);
    });
  });
  const ws3 = XLSX.utils.aoa_to_sheet(secRows);
  ws3["!cols"] = [12,8,18,6,28,18,10,10,8,30].map(w => ({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws3, "Section Summary");

  XLSX.writeFile(wb, filename + ".xlsx");
}

// ── Records table ─────────────────────────────────────────────────────────
function renderRecords() {
  const tbody = document.getElementById("records-body");
  if (records.length === 0) {
    tbody.innerHTML = '<tr><td colspan="8" class="empty-state">No records yet. Submit a completed form to see it here.</td></tr>';
    return;
  }
  tbody.innerHTML = records.slice().reverse().map((r, i) => {
    const pct = Math.round((r.tasksDone / r.tasksTotal) * 100);
    const complete = r.tasksDone === r.tasksTotal;
    const num = records.length - i;
    return `<tr>
      <td>${num}</td>
      <td>${r.date}</td>
      <td>${r.shift || "—"}</td>
      <td>${r.area || "—"}</td>
      <td>${r.supervisor || "—"} (${r.initials || "—"})</td>
      <td>${r.tasksDone}/${r.tasksTotal} (${pct}%)</td>
      <td><span class="badge ${complete ? 'badge-complete' : 'badge-partial'}">${complete ? "Complete" : "Incomplete"}</span></td>
      <td style="font-size:11px;color:var(--gray-400)">${new Date(r.submitted).toLocaleString()}</td>
    </tr>`;
  }).join("");
}

function clearRecords() {
  if (!confirm("Delete all " + records.length + " saved records? This cannot be undone.")) return;
  records = [];
  try { localStorage.removeItem("ssop_records"); } catch(e) {}
  renderRecords();
}

// ── Reset form ────────────────────────────────────────────────────────────
function resetForm() {
  if (!confirm("Reset all fields and checkboxes for a new shift entry?")) return;
  ["f-date","f-shift","f-area","f-sup","sig-initials","sig-time","sig-ca"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = id === "f-date" ? new Date().toISOString().split("T")[0] : "";
  });
  SECTIONS.forEach(sec => {
    const ti = document.getElementById("tech-" + sec.id);
    if (ti) ti.value = "";
    const ni = document.getElementById("notes-" + sec.id);
    if (ni) ni.value = "";
    sec.tasks.forEach((t, i) => {
      const key = sec.id + "_" + i;
      checks[key] = false;
      const cb = document.getElementById("chk-" + key);
      const lb = document.getElementById("lbl-" + key);
      if (cb) cb.checked = false;
      if (lb) lb.className = "task-text";
    });
    updateSection(sec.id, sec.tasks.length);
  });
  updateProgress();
  document.getElementById("success-banner").style.display = "none";
  window.scrollTo({top:0,behavior:"smooth"});
}

// ── Page nav ──────────────────────────────────────────────────────────────
function showPage(name) {
  document.querySelectorAll(".page").forEach(p => p.classList.remove("active"));
  document.querySelectorAll(".nav-tab").forEach(t => t.classList.remove("active"));
  document.getElementById("page-" + name).classList.add("active");
  document.querySelectorAll(".nav-tab")[name === "form" ? 0 : name === "records" ? 1 : 2].classList.add("active");
  if (name === "qr") generateQR();
  if (name === "records") renderRecords();
}

// ── QR Code ───────────────────────────────────────────────────────────────
let qrGenerated = false;
function generateQR() {
  if (qrGenerated) return;
  qrGenerated = true;
  const url = window.location.href.split("?")[0].split("#")[0];
  document.getElementById("qr-url-display").textContent = url;
  document.getElementById("qrcode").innerHTML = "";
  new QRCode(document.getElementById("qrcode"), {
    text: url, width: 220, height: 220,
    colorDark: "#1a2e4a", colorLight: "#ffffff",
    correctLevel: QRCode.CorrectLevel.H
  });
}

function downloadQR() {
  const canvas = document.querySelector("#qrcode canvas");
  if (!canvas) { alert("QR code not generated yet."); return; }
  const a = document.createElement("a");
  a.download = "SSOP_QR_Code.png";
  a.href = canvas.toDataURL("image/png");
  a.click();
}

// ── Init ──────────────────────────────────────────────────────────────────
document.getElementById("f-date").value = new Date().toISOString().split("T")[0];
buildSections();
renderRecords();
</script>
</body>
</html>
