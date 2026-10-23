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
// Entry Log columns (one row per beam):
// A Date | B Time | C User | D Machine | E Beam | F Cycle | G Tape | H Wt/Mtr | I Wt/Roll | J Action | K Device ID | L Role
const logNum = (v) => { if (v === null || v === undefined || v === '') return ''; const n = Number(v); return isNaN(n) ? safeText(v) : (n === 0 ? '' : n); };
async function log(meta, rows) {
  try {
    if (!rows.length) return;
    const titles = await G.sheetTitles(SHEET_ID);
    if (!titles.includes(LOG)) return;                       // never create tabs
    const t = nowParts(), m = meta || {};
    const out = rows.map((r) => [t.date, t.time, safeText(m.user || 'Unknown'), safeText(r.machine || ''), r.beam || '',
      logNum(r.cycle), safeText(r.tape || ''), logNum(r.weightMtr), logNum(r.weightRoll), safeText(r.action || ''), safeText(m.device || ''), 'Suryatex OS']);
    await G.append(SHEET_ID, at(LOG, 'A:L'), out, 'USER_ENTERED');
  } catch (e) { console.error('knitted log failed: ' + e.message); }
}
const hasData = (b) => [b.cycle, b.weightMtr, b.weightRoll].some((v) => Number(v)) || !!String(b.tape || '').trim();

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
  const [before] = await G.batchGet(SHEET_ID, [at(TRACKER, 'A1:K80')]);
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
  const name = MACHINE_NAMES[machineId] || machineId, rows = [];
  beams.forEach((b, i) => {
    const r = Object.assign({}, b || {}, { tape: tapes[i], machine: name, beam: i + 1 });
    if (hasData(r)) rows.push(Object.assign(r, { action: 'Save' }));
    else if (String(cell(before[L.rows[i] - 1], col(L.cols.tape))).trim()) rows.push({ machine: name, beam: i + 1, action: 'Cleared' });
  });
  await log(meta, rows);
  await refreshActivity(nowParts().date.slice(0, 7));
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
  await log(meta, [{ machine: 'Tape list', tape: clean, action: 'Add Tape' }]);
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
    await log(meta, [{ machine: 'Tape list', tape: clean, action: 'Remove Tape' }]);
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
  await log(meta, [{ machine: 'Tape list', tape: o + ' -> ' + n, action: 'Rename Tape' }]);
  return { tapeList: await tapeList(), updatedBeams: updated };
}

/* ---------- Machine activity & changeovers (read from Entry Log) ---------- */
// A machine is ACTIVE on a day when any beam's cycle or tape changed from its last saved value
// (or a beam was cleared). Changeover = one beam's cycle or tape changing. First sighting of a beam = baseline.
const NAME_TO_ID = Object.fromEntries(Object.entries(MACHINE_NAMES).map(([k, v]) => [v.toLowerCase(), k]));
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAY0 = Date.UTC(1899, 11, 30);
function serialDay(v) {                       // sheet date cell -> whole day number
  if (typeof v === 'number') return Math.floor(v);
  const s = String(v || '').trim(); let m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return Math.round((Date.UTC(+m[1], m[2] - 1, +m[3]) - DAY0) / 864e5);
  if ((m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/))) return Math.round((Date.UTC(+m[3], m[2] - 1, +m[1]) - DAY0) / 864e5);
  return null;
}
function dayFrac(v) {
  if (typeof v === 'number') return v % 1;
  const m = String(v || '').match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  return m ? (+m[1] * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400 : 0;
}
const ymd = (day) => new Date(DAY0 + day * 864e5).toISOString().slice(0, 10);
const cnum = (v) => { const n = Number(v); return v === '' || v == null || isNaN(n) || n === 0 ? null : n; };

const LATE_X = 1.5;                            // flag when the current gap is 1.5x the machine's own normal gap
const median = (xs) => { const v = xs.slice().sort((x, y) => x - y), n = v.length; return n ? (n % 2 ? v[(n - 1) / 2] : (v[n / 2 - 1] + v[n / 2]) / 2) : null; };
const r1 = (x) => x == null ? null : Math.round(x * 10) / 10;

async function computeActivity(month) {      // month = 'YYYY-MM'
  const [rows] = await G.batchGet(SHEET_ID, [at(LOG, 'A2:J')], 'serial');
  const events = [];
  for (const r of rows) {
    const id = NAME_TO_ID[String(cell(r, 3)).trim().toLowerCase()], beam = Number(cell(r, 4)), day = serialDay(cell(r, 0));
    if (!id || !(beam >= 1 && beam <= 4) || day == null) continue;             // skips old-format / team-site rows
    events.push({ id, beam, day, t: day + dayFrac(cell(r, 1)), cycle: cnum(cell(r, 5)), tape: String(cell(r, 6) || '').trim(), action: String(cell(r, 9) || '') });
  }
  events.sort((a, b) => a.t - b.t);
  const [Y, M] = month.split('-').map(Number);
  const first = Math.round((Date.UTC(Y, M - 1, 1) - DAY0) / 864e5), ndays = new Date(Date.UTC(Y, M, 0)).getUTCDate(), last = first + ndays - 1;
  const np = nowParts(), today = serialDay(np.date), now = today + dayFrac(np.time), start = events.length ? events[0].day : null;
  const isCurrent = today >= first && today <= last, until = isCurrent ? now : last + 1;    // look at history up to now / month end
  const out = {}; Object.keys(MACHINE_NAMES).forEach((id) => { out[id] = { cycleChanges: 0, tapeChanges: 0, stamps: {}, moments: [], seen: null }; });
  const lastSeen = {};
  for (const e of events) {
    if (e.t >= until) break;
    const k = e.id + ':' + e.beam, prev = lastSeen[k], inMonth = e.day >= first && e.day <= last, o = out[e.id];
    let changed = false;
    if (o.seen == null) o.seen = e.t;
    if (/clear/i.test(e.action)) changed = !!prev;
    else if (prev) {
      if (e.cycle != null && prev.cycle != null && e.cycle !== prev.cycle) { changed = true; if (inMonth) o.cycleChanges++; }
      if (e.tape && prev.tape && e.tape.toLowerCase() !== prev.tape.toLowerCase()) { changed = true; if (inMonth) o.tapeChanges++; }
    }
    if (changed) {
      const mm = o.moments; if (!mm.length || e.t - mm[mm.length - 1] > 1 / 1440) mm.push(e.t);   // one save = one change moment
      if (inMonth) (o.stamps[e.beam] = o.stamps[e.beam] || []).push(e.t);
    }
    lastSeen[k] = { cycle: e.cycle != null ? e.cycle : (prev && prev.cycle), tape: e.tape || (prev && prev.tape) || '' };
  }
  const fromDay = start == null ? null : Math.max(first, start), toDay = Math.min(last, today);
  const tracked = fromDay == null || toDay < fromDay ? 0 : toDay - fromDay + 1;
  const machines = Object.keys(MACHINE_NAMES).map((id) => {
    const o = out[id], mm = o.moments, gaps = [];
    for (let i = 1; i < mm.length; i++) gaps.push(mm[i] - mm[i - 1]);
    const normal = mm.length >= 3 ? median(gaps) : null;
    const lastT = mm.length ? mm[mm.length - 1] : null;
    const since = isCurrent && lastT != null ? now - lastT : null;
    const late = !!(normal && since != null && since > normal * LATE_X);
    // longest gap that touches this month (ongoing gap included for the current month)
    let longest = null;
    for (let i = 1; i < mm.length; i++) if (Math.floor(mm[i]) >= first && Math.floor(mm[i]) <= last) longest = Math.max(longest || 0, mm[i] - mm[i - 1]);
    if (since != null) longest = Math.max(longest || 0, since);
    // per-day state for the grid: change / running / late / no baseline
    const days = [];
    for (let d = 1; d <= ndays; d++) {
      const day = first + d - 1;
      if (fromDay == null || day < fromDay || day > toDay) { days.push(''); continue; }
      if (mm.some((t) => Math.floor(t) === day)) { days.push('change'); continue; }
      const end = day === today && isCurrent ? now : day + 1, before = mm.filter((t) => t < end);
      if (!before.length) { days.push(o.seen != null && o.seen < end ? 'run' : 'wait'); continue; }
      days.push(normal && end - before[before.length - 1] > normal * LATE_X ? 'late' : 'run');
    }
    const bgaps = [];
    Object.values(o.stamps).forEach((ts) => { for (let i = 1; i < ts.length; i++) bgaps.push(ts[i] - ts[i - 1]); });
    return { id, name: MACHINE_NAMES[id], changes: mm.filter((t) => Math.floor(t) >= first && Math.floor(t) <= last).length,
      cycleChanges: o.cycleChanges, tapeChanges: o.tapeChanges, changeovers: o.cycleChanges + o.tapeChanges,
      avgDays: bgaps.length ? r1(bgaps.reduce((a, b) => a + b, 0) / bgaps.length) : null,
      normalGap: r1(normal), historyChanges: mm.length, lastChange: lastT == null ? null : ymd(Math.floor(lastT)),
      daysSince: r1(since), longestGap: r1(longest), late, days };
  });
  return { month, label: MONTHS[M - 1] + ' ' + Y, days: ndays, current: isCurrent, lateFactor: LATE_X, firstTracked: start == null ? null : ymd(start),
    trackedFromDay: fromDay == null ? null : fromDay - first + 1, trackedToDay: toDay - first + 1, trackedDays: tracked, machines };
}

// keep one summary tab per month in the sheet ("Activity October 2026"), rewritten on each save
async function writeActivityTab(a) {
  const title = 'Activity ' + a.label;
  const props = await G.sheetProps(SHEET_ID);
  const ex = props.find((p) => p.title === title), need = 12 + a.days;
  if (ex && ex.gridProperties && ex.gridProperties.columnCount < need) {
    await G.format(SHEET_ID, [{ appendDimension: { sheetId: ex.sheetId, dimension: 'COLUMNS', length: need - ex.gridProperties.columnCount } }]);
  }
  if (!ex) {
    await G.format(SHEET_ID, [{ addSheet: { properties: { title, gridProperties: { rowCount: 20, columnCount: 12 + a.days, frozenRowCount: 2, frozenColumnCount: 1 } } } }]);
  }
  const head = ['Machine', 'Changes this month', 'Cycle changes', 'Tape changes', 'Total changeovers', 'Beam turnaround (days)', 'Normal gap (days)', 'Last change', 'Days since last change', 'Longest gap (days)', 'Status'];
  for (let d = 1; d <= a.days; d++) head.push(String(d));
  const word = { change: 'Change', run: '', late: 'Late', wait: '', '': '' };
  const body = a.machines.map((m) => [m.name, m.changes, m.cycleChanges, m.tapeChanges, m.changeovers, m.avgDays == null ? '' : m.avgDays,
    m.normalGap == null ? 'learning (' + m.historyChanges + '/3)' : m.normalGap, m.lastChange || '', m.daysSince == null ? '' : m.daysSince,
    m.longestGap == null ? '' : m.longestGap, m.late ? 'Running late' : (a.current ? 'OK' : '')].concat(m.days.map((x) => word[x])));
  const t = nowParts();
  await G.batchUpdate(SHEET_ID, [{ range: at(title, 'A1'), values: [['Machine Activity — ' + a.label + ' · updated ' + t.date + ' ' + t.time + ' · Change = a beam\'s cycle or tape changed · Late = gap longer than ' + a.lateFactor + 'x the machine\'s normal gap']] },
    { range: at(title, 'A2'), values: [head].concat(body) }], 'RAW');
}
async function refreshActivity(month) {
  try { const a = await computeActivity(month); await writeActivityTab(a); return a; }
  catch (e) { console.error('activity tab failed: ' + e.message); return null; }
}
async function activity(month, meta) {
  requireAdmin(meta);
  const cur = nowParts().date.slice(0, 7);
  month = /^\d{4}-\d{2}$/.test(String(month || '')) ? month : cur;
  const a = await computeActivity(month);
  if (month <= cur) { try { await writeActivityTab(a); } catch (e) { console.error('activity tab failed: ' + e.message); } }
  return a;
}

const ROUTES = {
  'knitted.getAll': () => getAll(),
  'knitted.saveBeams': (a, m) => saveBeams(a.machineId, a.beams, m),
  'knitted.verifyPin': (a) => verifyPin(a.pin),
  'knitted.addTape': (a, m) => addTape(a.name, m),
  'knitted.removeTape': (a, m) => removeTape(a.name, m),
  'knitted.renameTape': (a, m) => renameTape(a.oldName, a.newName, m),
  'knitted.activity': (a, m) => activity(a.month, m),
};

module.exports = { ROUTES, SHEET_ID };
