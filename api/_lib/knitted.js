// Knitted Production — direct Google Sheets version of the Hub's knitted.* functions.
// Same Tracker cells, same Product Description rules and same Entry Log columns as the Hub script
// (which mirrors the team's Knitted Beams Tracker). Role column = "Suryatex OS".
const G = require('./google');

const SHEET_ID = '1P-TyETzmarTEEQCsI_bQlpbWYadxlqdih657L9pLfvw';   // Knitted Production Tracker
const TRACKER = 'Tracker', PRODUCT = 'Product Description', LOG = 'Entry Log';
const TZ = 'Asia/Kolkata';

const LAYOUT = {
  comez:       { rows: [4, 7, 10, 13],   cols: { cycle: 'B', tape: 'C', weightMtr: 'D', weightRoll: 'E' } },
  mainMuller:  { rows: [4, 7, 10, 13],   cols: { cycle: 'H', tape: 'I', weightMtr: 'J', weightRoll: 'K' } },
  sideMuller:  { rows: [20, 23, 26, 29], cols: { cycle: 'B', tape: 'C', weightMtr: 'D', weightRoll: 'E' } },
  ky:          { rows: [20, 23, 26, 29], cols: { cycle: 'H', tape: 'I', weightMtr: 'J', weightRoll: 'K' } },
  tch1:        { rows: [36, 39, 42, 45], cols: { cycle: 'B', tape: 'C', weightMtr: 'D', weightRoll: 'E' } },
  tch2:        { rows: [36, 39, 42, 45], cols: { cycle: 'H', tape: 'I', weightMtr: 'J', weightRoll: 'K' } },
  fastMuller1: { rows: [52, 55, 58, 61], cols: { cycle: 'B', tape: 'C', weightMtr: 'D', weightRoll: 'E' } },
  fastMuller2: { rows: [52, 55, 58, 61], cols: { cycle: 'H', tape: 'I', weightMtr: 'J', weightRoll: 'K' } },
  victory1:    { rows: [68, 71, 74, 77], cols: { cycle: 'B', tape: 'C', weightMtr: 'D', weightRoll: 'E' } },
  victory2:    { rows: [68, 71, 74, 77], cols: { cycle: 'H', tape: 'I', weightMtr: 'J', weightRoll: 'K' } },
};

const col = (letter) => letter.charCodeAt(0) - 65;            // 'A' -> 0
const cell = (r, c) => (r && r[c] !== undefined && r[c] !== null) ? r[c] : '';
const norm = (s) => String(s == null ? '' : s).trim().toLowerCase();
const at = (tab, a1) => G.q(tab) + '!' + a1;

function nowParts() {
  const f = (o) => new Intl.DateTimeFormat('en-CA', Object.assign({ timeZone: TZ }, o)).format(new Date());
  return { date: f({ year: 'numeric', month: '2-digit', day: '2-digit' }),
           time: new Intl.DateTimeFormat('en-GB', { timeZone: TZ, hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }).format(new Date()),
           loaded: new Intl.DateTimeFormat('en-GB', { timeZone: TZ, day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit', hour12: false }).format(new Date()).replace(',', '') };
}

async function tapeList() {
  const [rows] = await G.batchGet(SHEET_ID, [at(PRODUCT, 'A2:A')]);
  return rows.map((r) => String(cell(r, 0))).filter((n) => n.trim().length > 0);   // exact text, trailing spaces kept
}
// whole column A from row 2, blanks included, plus the sheet's last used row (like getLastRow())
async function productColumn() {
  const [all] = await G.batchGet(SHEET_ID, [at(PRODUCT, 'A1:Z')]);
  const last = all.length;
  const values = [];
  for (let i = 1; i < last; i++) values.push(String(cell(all[i], 0)));
  return { last, values };
}

function resolveTape(input, list) {
  const clean = String(input == null ? '' : input).trim();
  if (!clean) return '';
  const m = list.find((t) => t.trim().toLowerCase() === clean.toLowerCase());
  if (!m) throw new Error('Tape "' + clean + '" is not in the Product Description list');
  return m;
}
// same rule as the Hub's kn_write_: blank = leave the cell alone; numbers stored as numbers
function writeVal(v) {
  if (v === null || v === undefined || v === '') return undefined;
  const n = Number(v);
  return isNaN(n) ? String(v) : n;
}
// stop text being read as a formula when appended with USER_ENTERED
const safeText = (s) => { s = String(s == null ? '' : s); return /^[=+\-@]/.test(s) ? "'" + s : s; };

const MACHINE_NAMES = {
  comez: 'Comez', mainMuller: 'Main Muller', sideMuller: 'Side Muller', ky: 'KY', tch1: 'TCH 1', tch2: 'TCH 2',
  fastMuller1: 'Fast Muller 1', fastMuller2: 'Fast Muller 2', victory1: 'Victory 1', victory2: 'Victory 2',
};
// Entry Log columns: A Date | B Time | C User | D Device | E Role | F Action | G Machine | H..W = 4 beams x (Cycle, Tape, Wt/Mtr, Wt/Roll)
const logNum = (v) => { if (v === null || v === undefined || v === '') return ''; const n = Number(v); return isNaN(n) ? safeText(v) : (n === 0 ? '' : n); };
function beamCols(beams) {
  const out = [];
  for (let i = 0; i < 4; i++) { const b = (beams && beams[i]) || {}; out.push(logNum(b.cycle), safeText(b.tape || ''), logNum(b.weightMtr), logNum(b.weightRoll)); }
  return out;
}
async function log(meta, action, machine, rest) {
  try {
    const titles = await G.sheetTitles(SHEET_ID);
    if (!titles.includes(LOG)) return;                       // never create tabs
    const t = nowParts(), m = meta || {};
    const row = [t.date, t.time, safeText(m.user || 'Unknown'), safeText(m.device || 'Unknown'), 'Suryatex OS', safeText(action || ''), safeText(machine || '')].concat(rest || []);
    await G.append(SHEET_ID, at(LOG, 'A:W'), [row], 'USER_ENTERED');
  } catch (e) { console.error('knitted log failed: ' + e.message); }
}

async function getAll() {
  const [grid, prod] = await G.batchGet(SHEET_ID, [at(TRACKER, 'A1:K80'), at(PRODUCT, 'A2:A')]);
  const cycleData = {};
  for (const id of Object.keys(LAYOUT)) {
    const L = LAYOUT[id];
    cycleData[id] = L.rows.map((row) => {
      const r = grid[row - 1];
      return { cycle: cell(r, col(L.cols.cycle)), tape: cell(r, col(L.cols.tape)), weightMtr: cell(r, col(L.cols.weightMtr)), weightRoll: cell(r, col(L.cols.weightRoll)) };
    });
  }
  return { cycleData, tapeList: prod.map((r) => String(cell(r, 0))).filter((n) => n.trim().length > 0), loaded: nowParts().loaded };
}

async function saveBeams(machineId, beams, meta) {
  const L = LAYOUT[machineId];
  if (!L) throw new Error('Unknown machine: ' + machineId);
  if (!Array.isArray(beams) || beams.length !== L.rows.length) throw new Error('Expected ' + L.rows.length + ' beams for ' + machineId);
  const list = await tapeList();
  const tapes = beams.map((b) => resolveTape(b && b.tape, list));   // fail before writing anything
  const data = [];
  L.rows.forEach((row, i) => {
    const b = beams[i] || {};
    const put = (c, v) => { if (v !== undefined) data.push({ range: at(TRACKER, L.cols[c] + row), values: [[v]] }); };
    put('cycle', writeVal(b.cycle));
    data.push({ range: at(TRACKER, L.cols.tape + row), values: [[tapes[i]]] });   // always written: blank clears the cell (Clear button)
    put('weightMtr', writeVal(b.weightMtr));
    put('weightRoll', writeVal(b.weightRoll));
  });
  await G.batchUpdate(SHEET_ID, data, 'RAW');
  await log(meta, 'Save Beams', MACHINE_NAMES[machineId] || machineId, beamCols(beams.map((b, i) => Object.assign({}, b, { tape: tapes[i] }))));
  return true;
}

function verifyPin(pin) { const p = process.env.KN_ADMIN_PIN || '1408'; return String(pin == null ? '' : pin).trim() === p; }
function requireAdmin(meta) { if (!meta || !verifyPin(meta.pin)) throw new Error('Admin PIN required'); }

async function addTape(name, meta) {
  requireAdmin(meta);
  const clean = String(name == null ? '' : name).trim();
  if (!clean) throw new Error('Tape name cannot be empty');
  const { last } = await productColumn();
  const list = await tapeList();
  if (list.some((t) => norm(t) === norm(clean))) return list;
  await G.batchUpdate(SHEET_ID, [{ range: at(PRODUCT, 'A' + (last + 1)), values: [[clean]] }], 'RAW');
  await log(meta, 'Add Tape', 'Tape list', [safeText(clean)]);
  return tapeList();
}

// shift names up instead of deleting the row, so the Tracker dropdown range never shrinks
async function removeTape(name, meta) {
  requireAdmin(meta);
  const clean = String(name == null ? '' : name).trim();
  if (!clean) throw new Error('Tape name cannot be empty');
  const { last, values } = await productColumn();
  if (last < 2) return tapeList();
  const idx = values.findIndex((v) => norm(v) === norm(clean));
  if (idx !== -1) {
    values.splice(idx, 1); values.push('');
    await G.batchUpdate(SHEET_ID, [{ range: at(PRODUCT, 'A2:A' + last), values: values.map((v) => [v]) }], 'RAW');
    await log(meta, 'Remove Tape', 'Tape list', [safeText(clean)]);
  }
  return tapeList();
}

// rename in the list AND in every Tracker beam that used the old name
async function renameTape(oldName, newName, meta) {
  requireAdmin(meta);
  const o = String(oldName == null ? '' : oldName).trim(), n = String(newName == null ? '' : newName).trim();
  if (!o || !n) throw new Error('Tape name cannot be empty');
  const { last, values } = await productColumn();
  if (last < 2) throw new Error('Tape "' + o + '" was not found');
  const idx = values.findIndex((v) => norm(v) === norm(o));
  if (idx === -1) throw new Error('Tape "' + o + '" was not found');
  if (values.some((v, i) => i !== idx && norm(v) === norm(n))) throw new Error('A tape named "' + n + '" already exists');
  const [grid] = await G.batchGet(SHEET_ID, [at(TRACKER, 'A1:K80')]);
  const data = [{ range: at(PRODUCT, 'A' + (idx + 2)), values: [[n]] }];
  let updated = 0;
  for (const id of Object.keys(LAYOUT)) {
    const L = LAYOUT[id];
    for (const row of L.rows) {
      if (norm(cell(grid[row - 1], col(L.cols.tape))) === norm(o)) { data.push({ range: at(TRACKER, L.cols.tape + row), values: [[n]] }); updated++; }
    }
  }
  await G.batchUpdate(SHEET_ID, data, 'RAW');
  await log(meta, 'Rename Tape', 'Tape list', [safeText(o + ' -> ' + n)]);
  return { tapeList: await tapeList(), updatedBeams: updated };
}

const ROUTES = {
  'knitted.getAll': () => getAll(),
  'knitted.saveBeams': (a, m) => saveBeams(a.machineId, a.beams, m),
  'knitted.verifyPin': (a) => verifyPin(a.pin),
  'knitted.addTape': (a, m) => addTape(a.name, m),
  'knitted.removeTape': (a, m) => removeTape(a.name, m),
  'knitted.renameTape': (a, m) => renameTape(a.oldName, a.newName, m),
};

module.exports = { ROUTES, SHEET_ID };
