/**
 * CCTV Daily Report Register (ICTU) - receives reports from "CCTV Daily Report.html"
 * and adds them to the "Daily Reports" and "Incidents" tabs.
 *
 * SETUP (one time, in the CCTV Daily Report Register (ICTU) Google Sheet,
 * signed in as salagfacility@gmail.com):
 *   1. Extensions > Apps Script. Delete what is there and paste this whole file. Save.
 *   2. Pick "setup" in the function list at the top and click Run. Allow the permissions.
 *   3. Deploy > New deployment > type: Web app.
 *        Execute as: Me      Who has access: Anyone
 *      Click Deploy and copy the Web app URL (ends in /exec).
 *   4. In CCTV Daily Report.html open "Google Sheet connection", paste the URL, Save settings.
 *
 * UPDATING an existing deployment (keeps the same URL):
 *   Deploy > Manage deployments > pencil icon > Version: New version > Deploy.
 */

const TZ = 'Asia/Manila';
const REPORTS = 'Daily Reports';
const INCIDENTS = 'Incidents';
const R_HEADERS = ['Report ID', 'Date', 'Shift', 'Operator on Duty', 'Position',
  'IP Installed', 'IP Working', 'IP Not Working (ID / location)', 'NVR Recording', 'NVR Date/Time Correct', 'NVR Days Stored',
  'Analog Installed', 'Analog Working', 'Analog Not Working (ID / location)', 'DVR Recording', 'DVR Date/Time Correct', 'DVR Days Stored',
  'Incident Status', 'No. of Incidents', 'Handover / Remarks', 'Noted By (DPO-Designate)', 'Saved At'];
const I_HEADERS = ['Report ID', 'Date', 'Shift', 'Incident No.', 'Time', 'Camera / Location',
  'What Happened', 'Clip Preserved', 'Action Taken', 'Operator on Duty'];

function doPost(e) {
  try {
    const r = JSON.parse(e.postData.contents).report;
    if (!r || !r.date) return reply_({ ok: false, error: 'No report received.' });

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);  // two people saving at once won't overwrite each other
    try {
      const rs = sheet_(REPORTS), is = sheet_(INCIDENTS);
      const last = lastRow_(rs);
      // Same date and shift already saved? Ask the page to confirm first.
      if (!r.force && last > 1) {
        const prev = rs.getRange(2, 1, last - 1, 3).getValues();
        const iso = d => d instanceof Date ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : String(d);
        const hit = prev.find(p => iso(p[1]) === r.date && String(p[2]) === (r.shift || ''));
        if (hit) return reply_({ ok: false, duplicate: hit[0] });
      }
      const id = nextId_(rs, last, r.date);
      const ip = r.ip || {}, an = r.an || {}, incs = r.incidents || [];
      rs.getRange(last + 1, 1, 1, R_HEADERS.length).setValues([[
        id, toDate_(r.date), r.shift || '', r.operator || '', r.position || '',
        num_(ip.total), num_(ip.up), ip.down || '', ip.rec || '', ip.clock || '', num_(ip.days),
        num_(an.total), num_(an.up), an.down || '', an.rec || '', an.clock || '', num_(an.days),
        incs.length ? 'Incident observed' : 'No incident', incs.length, r.notes || '', r.notedBy || '', new Date()
      ]]);
      rs.getRange(last + 1, 1).setNumberFormat('@');
      rs.getRange(last + 1, 2).setNumberFormat('mmm d, yyyy');
      rs.getRange(last + 1, 22).setNumberFormat('mmm d, yyyy h:mm am/pm');
      if (incs.length) {
        const start = lastRow_(is) + 1;
        is.getRange(start, 1, incs.length, I_HEADERS.length).setValues(incs.map((x, i) => [
          id, toDate_(r.date), r.shift || '', i + 1, x.time || '', x.cam || '', x.what || '', x.kept || '', x.action || '', r.operator || ''
        ]));
        is.getRange(start, 1, incs.length, 1).setNumberFormat('@');
        is.getRange(start, 2, incs.length, 1).setNumberFormat('mmm d, yyyy');
        is.getRange(start, 5, incs.length, 1).setNumberFormat('@');
      }
      SpreadsheetApp.flush();
      return reply_({ ok: true, id: id, row: last + 1, incidents: incs.length });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return reply_({ ok: false, error: String(err) });
  }
}

// Lets you open the /exec URL in a browser to check the script is live.
function doGet() {
  return reply_({ ok: true, message: 'CCTV Daily Report Register script is running.' });
}

// Run once from the editor. Makes sure both tabs and header rows exist.
function setup() {
  [[REPORTS, R_HEADERS], [INCIDENTS, I_HEADERS]].forEach(([name, h]) => {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getRange(1, 1).getValue() !== h[0]) sh.getRange(1, 1, 1, h.length).setValues([h]);
    sh.setFrozenRows(1);
  });
  SpreadsheetApp.getActive().setSpreadsheetTimeZone(TZ);
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Tab "' + name + '" not found. Run setup first.');
  return sh;
}

// Last row with something in column A (ignores empty formatted rows).
function lastRow_(sh) {
  const vals = sh.getRange(1, 1, sh.getMaxRows(), 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) if (vals[i][0] !== '') return i + 1;
  return 1;
}

// CCTV-2026-10-15-01, -02 ... for each report saved on the same date.
function nextId_(sh, last, iso) {
  const base = 'CCTV-' + iso + '-';
  let n = 0;
  if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(v => {
    const s = String(v[0]);
    if (s.indexOf(base) === 0) n = Math.max(n, +s.slice(base.length) || 0);
  });
  return base + String(n + 1).padStart(2, '0');
}

function num_(v) { return v === '' || v == null || isNaN(+v) ? (v || '') : +v; }

function toDate_(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : (s || '');
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
