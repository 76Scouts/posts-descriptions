/**
 * 76ο Σύστημα — Σημειώσεις Δημοσιεύσεων · backend (Apps Script Web App)
 * Πηγή αλήθειας: Google Sheet «ΔΗΜΟΣΙΕΥΣΕΙΣ_STATE», φύλλο «Πρώτη Ύλη».
 * Οι «Δράσεις στο site» διαβάζονται απευθείας από το 76.life μέσα από τη σελίδα — δεν περνούν από εδώ.
 *
 * API (GET, JSON):
 *   ?action=list                         -> { ok, raw:[...] }
 *   ?action=add&<πεδία>                  -> { ok, id }
 *   ?action=update&ID=..&<πεδία>         -> { ok }
 *   ?action=delete&ID=..&pin=....        -> { ok }   (χρειάζεται PIN διαγραφής)
 *
 * Μοντέλο append (μία γραμμή ανά σημείωση, μοναδικό ID) → ταυτόχρονες εγγραφές δεν σβήνουν η μία την άλλη.
 * Ημερομηνίες αποθηκεύονται ως ΚΕΙΜΕΝΟ (YYYY-MM-DD) ώστε το Sheets να μην τις μετατρέπει σε Date.
 */

const SHEET_ID   = '18i4X-3NGVDkJscf4jzDOagyzkGd9y3dCa_B5W0Os1ck'; // «ΔΗΜΟΣΙΕΥΣΕΙΣ_STATE» (as@76.life)
const DELETE_PIN = '476';                                          // PIN μόνο για διαγραφή (η αποθήκευση είναι ελεύθερη)

const RAW_SHEET   = 'Πρώτη Ύλη';
const RAW_HEADERS = ['ID', 'Καταχώρηση', 'Ημερομηνία', 'Έως', 'Τμήμα', 'Τίτλος', 'Κείμενο', 'Συντάκτης'];
const DATE_COLS   = ['Ημερομηνία', 'Έως'];

function ss_() { return SpreadsheetApp.getActive() || SpreadsheetApp.openById(SHEET_ID); }

function sheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(RAW_SHEET);
  if (!sh) {
    sh = ss.insertSheet(RAW_SHEET);
    sh.getRange(1, 1, 1, RAW_HEADERS.length).setValues([RAW_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Τρέξε ΜΙΑ φορά: φτιάχνει το φύλλο, βάζει τις στήλες ημερομηνίας σε «κείμενο» και διορθώνει τις παλιές τιμές. */
function initSheet() {
  const sh = sheet_();
  const tz = Session.getScriptTimeZone();
  DATE_COLS.forEach(h => {
    const c = RAW_HEADERS.indexOf(h) + 1;
    const n = Math.max(sh.getMaxRows() - 1, 1);
    const rng = sh.getRange(2, c, n, 1);
    const vals = rng.getValues().map(r => [r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : r[0]]);
    rng.setNumberFormat('@').setValues(vals);
  });
}

function cell_(v, h, tz) {
  if (v === undefined || v === null) return '';
  if (v instanceof Date) return DATE_COLS.indexOf(h) >= 0 ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v.toISOString();
  return String(v);
}

function read_() {
  const sh = sheet_();
  const tz = Session.getScriptTimeZone();
  const vals = sh.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (!String(r[0] || '').trim()) continue;
    const o = {};
    RAW_HEADERS.forEach((h, j) => o[h] = cell_(r[j], h, tz));
    rows.push(o);
  }
  return rows;
}

function addRow_(obj) {
  const sh = sheet_();
  const id = 'id' + Date.now() + Math.floor(Math.random() * 1000);
  obj.ID = id;
  obj['Καταχώρηση'] = new Date().toISOString();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const row = sh.getLastRow() + 1;
    const rng = sh.getRange(row, 1, 1, RAW_HEADERS.length);
    rng.setNumberFormat('@');                                   // όλα ως κείμενο → καμία αυτόματη μετατροπή
    rng.setValues([RAW_HEADERS.map(h => (obj[h] !== undefined && obj[h] !== null) ? String(obj[h]) : '')]);
  } finally { lock.releaseLock(); }
  return id;
}

function findRow_(sh, id) {
  const ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (let i = 1; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 1;
  return 0;
}

function updateRow_(id, obj) {
  const sh = sheet_();
  const row = findRow_(sh, id);
  if (!row) return false;
  RAW_HEADERS.forEach((h, j) => {
    if (h === 'ID' || h === 'Καταχώρηση' || obj[h] === undefined) return;
    sh.getRange(row, j + 1).setNumberFormat('@').setValue(String(obj[h]));
  });
  return true;
}

function deleteRow_(id) {
  const sh = sheet_();
  const row = findRow_(sh, id);
  if (!row) return false;
  sh.deleteRow(row);
  return true;
}

function out_(obj, cb) {
  const t = JSON.stringify(obj);
  if (cb) return ContentService.createTextOutput(cb + '(' + t + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.JSON);
}

function handle_(e) {
  const p = (e && e.parameter) || {};
  const cb = p.callback || '';
  const action = p.action || 'list';
  try {
    if (action === 'list') return out_({ ok: true, raw: read_(), ts: new Date().toISOString() }, cb);
    if (action === 'add' || action === 'update') {
      const obj = {};
      RAW_HEADERS.forEach(h => { if (p[h] !== undefined) obj[h] = p[h]; });
      if (action === 'add') return out_({ ok: true, id: addRow_(obj) }, cb);
      return out_({ ok: updateRow_(p.ID || p.id, obj) }, cb);
    }
    if (action === 'delete') {
      if (String(p.pin || '').trim() !== DELETE_PIN) return out_({ ok: false, error: 'pin' }, cb);
      return out_({ ok: deleteRow_(p.ID || p.id) }, cb);
    }
    return out_({ ok: false, error: 'unknown action: ' + action }, cb);
  } catch (err) {
    return out_({ ok: false, error: String(err) }, cb);
  }
}

function doGet(e)  { return handle_(e); }
function doPost(e) {
  if (e && e.postData && e.postData.contents) {
    try { e.parameter = Object.assign({}, e.parameter, JSON.parse(e.postData.contents)); } catch (_) {}
  }
  return handle_(e);
}
