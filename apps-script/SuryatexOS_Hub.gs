/**
 * Suryatex OS Hub — one standalone script that feeds Suryatex OS (suryatexos.vercel.app).
 *
 * This is a SEPARATE Apps Script project. It is not attached to any sheet and does not touch
 * any of your team's existing scripts (Code.gs / Index) or their Google Sites deployments.
 * It opens each function's Google Sheet by its ID.
 *
 * Functions so far:
 *   knitted.*  — Knitted Production Tracker (same cells and Entry Log format as the team's Knitted Beams Tracker)
 *
 * Setup (once):
 *   1. Go to script.new, name the project "Suryatex OS Hub", replace Code.gs with this file, Save.
 *   2. Pick "hubTest" in the function menu and press Run. Allow the permissions Google asks for.
 *      It only READS and logs what it found.
 *   3. Deploy -> New deployment -> type: Web app -> Execute as: Me, Who has access: Anyone -> Deploy.
 *   4. Send the Web app URL (ends in /exec).
 * Later code changes: Deploy -> Manage deployments -> edit -> Version: New version (keeps the URL).
 *
 * Only requests carrying HUB_TOKEN (kept on the Suryatex OS server) are answered.
 */

var HUB_TOKEN = 'Bsq5KfpKmSZk6_JAwVAssOixDwbYmugb';   // must match HUB_SCRIPT_TOKEN on Vercel
var TZ = 'Asia/Kolkata';

var SHEET_IDS = {
  knitted: '1P-TyETzmarTEEQCsI_bQlpbWYadxlqdih657L9pLfvw'   // Knitted Production Tracker
};

function doGet() { return hubJson_({ ok: true, service: 'Suryatex OS Hub' }); }

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (x) {}
  if (body.token !== HUB_TOKEN) return hubJson_({ ok: false, error: 'Unauthorized' });
  var fn = HUB_ROUTES[body.fn];
  if (!fn) return hubJson_({ ok: false, error: 'Unknown function' });
  try {
    return hubJson_({ ok: true, data: fn(body.args || {}, body.meta || {}) });
  } catch (err) {
    return hubJson_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

var HUB_ROUTES = {
  'knitted.getAll':     function (a)    { return kn_getAll_(); },
  'knitted.saveBeams':  function (a, m) { return kn_saveBeams_(a.machineId, a.beams, m); },
  'knitted.verifyPin':  function (a)    { return kn_verifyPin_(a.pin); },
  'knitted.addTape':    function (a, m) { return kn_addTape_(a.name, m); },
  'knitted.removeTape': function (a, m) { return kn_removeTape_(a.name, m); },
  'knitted.renameTape': function (a, m) { return kn_renameTape_(a.oldName, a.newName, m); }
};

function hubJson_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Run from the editor once: asks for permission and checks each sheet can be read. Reads only.
function hubTest() {
  var d = kn_getAll_();
  var beams = 0;
  Object.keys(d.cycleData).forEach(function (m) { beams += d.cycleData[m].length; });
  Logger.log('Knitted: ' + Object.keys(d.cycleData).length + ' machines, ' + beams + ' beams, ' + d.tapeList.length + ' tapes');
}

/* ============================================================================
 * KNITTED PRODUCTION — mirrors the team's Knitted Beams Tracker exactly:
 * same Tracker cells, same Product Description list rules, same Entry Log columns.
 * ========================================================================== */

var KN_TRACKER = 'Tracker';
var KN_PRODUCT = 'Product Description';
var KN_LOG = 'Entry Log';
var KN_ADMIN_PIN = '1408';            // same PIN the team already uses for Manage Tapes

var KN_LAYOUT = {
  comez:        { rows: [4, 7, 10, 13],   cols: { cycle: 'B', tape: 'C', weightMtr: 'D', weightRoll: 'E' } },
  mainMuller:   { rows: [4, 7, 10, 13],   cols: { cycle: 'H', tape: 'I', weightMtr: 'J', weightRoll: 'K' } },
  sideMuller:   { rows: [20, 23, 26, 29], cols: { cycle: 'B', tape: 'C', weightMtr: 'D', weightRoll: 'E' } },
  ky:           { rows: [20, 23, 26, 29], cols: { cycle: 'H', tape: 'I', weightMtr: 'J', weightRoll: 'K' } },
  tch1:         { rows: [36, 39, 42, 45], cols: { cycle: 'B', tape: 'C', weightMtr: 'D', weightRoll: 'E' } },
  tch2:         { rows: [36, 39, 42, 45], cols: { cycle: 'H', tape: 'I', weightMtr: 'J', weightRoll: 'K' } },
  fastMuller1:  { rows: [52, 55, 58, 61], cols: { cycle: 'B', tape: 'C', weightMtr: 'D', weightRoll: 'E' } },
  fastMuller2:  { rows: [52, 55, 58, 61], cols: { cycle: 'H', tape: 'I', weightMtr: 'J', weightRoll: 'K' } },
  victory1:     { rows: [68, 71, 74, 77], cols: { cycle: 'B', tape: 'C', weightMtr: 'D', weightRoll: 'E' } },
  victory2:     { rows: [68, 71, 74, 77], cols: { cycle: 'H', tape: 'I', weightMtr: 'J', weightRoll: 'K' } }
};

function kn_ss_() { return SpreadsheetApp.openById(SHEET_IDS.knitted); }
function kn_colNum_(letter) { return letter.charCodeAt(0) - 64; }

function kn_getAll_() {
  var ss = kn_ss_();
  return { cycleData: kn_cycleData_(ss), tapeList: kn_tapeList_(ss),
           loaded: Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm') };
}

// One read of the Tracker block instead of 160 single-cell reads; same values as getValue().
function kn_cycleData_(ss) {
  var sheet = ss.getSheetByName(KN_TRACKER);
  if (!sheet) throw new Error('Tracker sheet not found');
  var grid = sheet.getRange(1, 1, 80, 11).getValues();       // A1:K80 covers every machine block
  var out = {};
  Object.keys(KN_LAYOUT).forEach(function (id) {
    var L = KN_LAYOUT[id];
    out[id] = L.rows.map(function (row) {
      var r = grid[row - 1];
      return { cycle: r[kn_colNum_(L.cols.cycle) - 1], tape: r[kn_colNum_(L.cols.tape) - 1],
               weightMtr: r[kn_colNum_(L.cols.weightMtr) - 1], weightRoll: r[kn_colNum_(L.cols.weightRoll) - 1] };
    });
  });
  return out;
}

// Exact cell text (trailing spaces kept) — the Tracker dropdown validates against these strings.
function kn_tapeList_(ss) {
  var sheet = ss.getSheetByName(KN_PRODUCT);
  if (!sheet) return [];
  var last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, 1).getValues()
    .map(function (r) { return (r[0] || '').toString(); })
    .filter(function (n) { return n.trim().length > 0; });
}

function kn_resolveTape_(input, list) {
  var clean = (input || '').toString().trim();
  if (!clean) return '';
  for (var i = 0; i < list.length; i++) if (list[i].trim().toLowerCase() === clean.toLowerCase()) return list[i];
  throw new Error('Tape "' + clean + '" is not in the Product Description list');
}

// Same rule as the team's writeNumericOrText: blank values are skipped (cell left as it was).
function kn_write_(sheet, a1, value) {
  if (value === null || value === undefined || value === '') return;
  var num = Number(value);
  sheet.getRange(a1).setValue(isNaN(num) ? value : num);
}

function kn_saveBeams_(machineId, beams, meta) {
  var L = KN_LAYOUT[machineId];
  if (!L) throw new Error('Unknown machine: ' + machineId);
  if (!beams || beams.length !== L.rows.length) throw new Error('Expected ' + L.rows.length + ' beams for ' + machineId);
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var ss = kn_ss_(), sheet = ss.getSheetByName(KN_TRACKER);
    var list = kn_tapeList_(ss);
    var tapes = beams.map(function (b) { return kn_resolveTape_(b && b.tape, list); });   // fail before writing anything
    L.rows.forEach(function (row, i) {
      var b = beams[i] || {};
      kn_write_(sheet, L.cols.cycle + row, b.cycle);
      kn_write_(sheet, L.cols.tape + row, tapes[i]);
      kn_write_(sheet, L.cols.weightMtr + row, b.weightMtr);
      kn_write_(sheet, L.cols.weightRoll + row, b.weightRoll);
    });
    kn_log_(ss, meta, 'Save Beams: ' + machineId, JSON.stringify(beams));
    return true;
  } finally { lock.releaseLock(); }
}

function kn_verifyPin_(pin) { return (pin || '').toString().trim() === KN_ADMIN_PIN; }
function kn_requireAdmin_(meta) { if (!meta || !kn_verifyPin_(meta.pin)) throw new Error('Admin PIN required'); }

// Entry Log: same 7 columns as the team's app. Role is "Suryatex OS" so entries made here are easy to tell apart.
function kn_log_(ss, meta, action, details) {
  try {
    var sheet = ss.getSheetByName(KN_LOG);
    if (!sheet) return;                      // never create tabs; the team's app owns the sheet layout
    var now = new Date(), m = meta || {};
    sheet.appendRow([
      Utilities.formatDate(now, TZ, 'yyyy-MM-dd'), Utilities.formatDate(now, TZ, 'HH:mm:ss'),
      m.user || 'Unknown', m.device || 'Unknown', 'Suryatex OS', action || '', details || ''
    ]);
  } catch (err) { console.error('log failed: ' + err); }
}

function kn_addTape_(name, meta) {
  kn_requireAdmin_(meta);
  var clean = (name || '').toString().trim();
  if (!clean) throw new Error('Tape name cannot be empty');
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var ss = kn_ss_(), sheet = ss.getSheetByName(KN_PRODUCT);
    if (!sheet) throw new Error('Product Description sheet not found');
    var list = kn_tapeList_(ss);
    if (list.some(function (t) { return t.trim().toLowerCase() === clean.toLowerCase(); })) return list;
    sheet.getRange(sheet.getLastRow() + 1, 1).setValue(clean);
    kn_log_(ss, meta, 'Add Tape', clean);
    return kn_tapeList_(ss);
  } finally { lock.releaseLock(); }
}

// Shifts names up instead of deleting the row, so the Tracker dropdown range never shrinks.
function kn_removeTape_(name, meta) {
  kn_requireAdmin_(meta);
  var clean = (name || '').toString().trim();
  if (!clean) throw new Error('Tape name cannot be empty');
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var ss = kn_ss_(), sheet = ss.getSheetByName(KN_PRODUCT);
    if (!sheet) throw new Error('Product Description sheet not found');
    var last = sheet.getLastRow();
    if (last < 2) return kn_tapeList_(ss);
    var range = sheet.getRange(2, 1, last - 1, 1);
    var values = range.getValues().map(function (r) { return r[0]; });
    var idx = values.findIndex(function (v) { return (v || '').toString().trim().toLowerCase() === clean.toLowerCase(); });
    if (idx !== -1) {
      values.splice(idx, 1); values.push('');
      range.setValues(values.map(function (v) { return [v]; }));
      kn_log_(ss, meta, 'Remove Tape', clean);
    }
    return kn_tapeList_(ss);
  } finally { lock.releaseLock(); }
}

// Renames in the list AND in every Tracker beam that used the old name.
function kn_renameTape_(oldName, newName, meta) {
  kn_requireAdmin_(meta);
  var o = (oldName || '').toString().trim(), n = (newName || '').toString().trim();
  if (!o || !n) throw new Error('Tape name cannot be empty');
  var norm = function (s) { return (s || '').toString().trim().toLowerCase(); };
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var ss = kn_ss_(), prod = ss.getSheetByName(KN_PRODUCT);
    if (!prod) throw new Error('Product Description sheet not found');
    var last = prod.getLastRow();
    if (last < 2) throw new Error('Tape "' + o + '" was not found');
    var values = prod.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return (r[0] || '').toString(); });
    var idx = values.findIndex(function (v) { return norm(v) === norm(o); });
    if (idx === -1) throw new Error('Tape "' + o + '" was not found');
    if (values.some(function (v, i) { return i !== idx && norm(v) === norm(n); })) throw new Error('A tape named "' + n + '" already exists');
    prod.getRange(idx + 2, 1).setValue(n);
    var tracker = ss.getSheetByName(KN_TRACKER), updated = 0;
    Object.keys(KN_LAYOUT).forEach(function (id) {
      var L = KN_LAYOUT[id];
      L.rows.forEach(function (row) {
        var cell = tracker.getRange(L.cols.tape + row);
        if (norm(cell.getValue()) === norm(o)) { cell.setValue(n); updated++; }
      });
    });
    kn_log_(ss, meta, 'Rename Tape', o + ' -> ' + n + ' (' + updated + ' beams updated)');
    return { tapeList: kn_tapeList_(ss), updatedBeams: updated };
  } finally { lock.releaseLock(); }
}
