// Customer Touch Point Tracker (Sales) — direct Google Sheets access, same sheet and layout the team's Apps Script uses.
// Env: TOUCH_SHEET_ID (default = the Snapweaves_Touch_Point_Tracker_FINAL workbook)
//      TOUCH_ROLLOVER=1 turns on the automatic month-end archive + reset (leave OFF while the old Apps Script app is still in use,
//      because the two would disagree about "what month is it" around midnight if their time zones differ), TOUCH_TZ (default Asia/Kolkata).
const G = require('./google');

const SHEET_ID = () => process.env.TOUCH_SHEET_ID || '18Ko-4kyw_BhORLGIZ_zTWe3M0atUS9LpIA_QQCDbAtU';
const ARCHIVE = 'Archive - Previous Months', ARCHIVED = 'Archived Customers';
const FIRST = 4, TP = 5, MAXROW = 1000;
// 1-indexed columns, exactly as in the sheet (same table the Apps Script's CLASS_CONFIG holds)
const CFG = {
  A: { sheet: 'Class A Customers', num: 1, name: 2, phone: 3, state: 4, split: false, order: 5, tp: 6, comp: 11, status: 12, notes: 13 },
  B: { sheet: 'Class B Customers', num: 1, name: 2, phone: 3, state: 4, split: true, total: 5, avg: 6, tp: 7, comp: 12, status: 13, notes: 14 },
};
const TPNAMES = ['Order Follow-up', 'Payment / Dispatch / Quotation', 'New Pattern / Catalog', 'Relationship Call / Visit', 'Sample / Quality Feedback'];

const L = (n) => { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
const cell = (cfg, col, row) => G.q(cfg.sheet) + '!' + L(col) + row;
const cfgOf = (cls) => { const c = CFG[cls]; if (!c) throw new Error('Unknown class'); return c; };
const str = (v) => (v === null || v === undefined ? '' : String(v));
const clip = (v, n) => str(v).trim().slice(0, n);
const num = (v) => { if (v === '' || v === null || v === undefined) return ''; const n = Number(v); return Number.isFinite(n) && n >= 0 ? n : ''; };
const statusOf = (done) => (done === 0 ? 'Not Started' : done >= 2 ? 'Completed' : 'On Track');

function monthLabel(d, tz) { return new Intl.DateTimeFormat('en-US', { timeZone: tz, month: 'long', year: 'numeric' }).format(d); }
// The "Active month" cell can hold text ("October 2026") or, if Sheets converted it, a date serial number.
function normMonth(v) {
  if (typeof v === 'number') return monthLabel(new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000), 'UTC');
  return str(v).trim();
}
const fresh = (cls, rows, month) => ({ activeMonth: month, rows, hasSplitValue: CFG[cls].split });

function parseRows(cls, values) {
  const c = CFG[cls], out = [];
  values.forEach((r, i) => {
    const name = str(r[c.name - 1]).trim(); if (!name) return;
    const tp = []; for (let t = 0; t < TP; t++) tp.push(str(r[c.tp - 1 + t]).trim() === 'Done' ? 1 : 0);
    const done = tp.reduce((a, b) => a + b, 0);
    const rec = { rowIndex: FIRST + i, name, phone: str(r[c.phone - 1]).trim(), state: str(r[c.state - 1]).trim(), tp,
      completion: done / TP, status: statusOf(done), notes: str(r[c.notes - 1]) };
    if (c.split) { rec.totalValue = r[c.total - 1] === undefined ? '' : r[c.total - 1]; rec.avgOrderValue = r[c.avg - 1] === undefined ? '' : r[c.avg - 1]; }
    else rec.orderValue = r[c.order - 1] === undefined ? '' : r[c.order - 1];
    out.push(rec);
  });
  return out;
}

async function getAll() {
  await rollover();
  const [mA, mB, dA, dB] = await G.batchGet(SHEET_ID(), [
    G.q(CFG.A.sheet) + '!B2', G.q(CFG.B.sheet) + '!B2',
    G.q(CFG.A.sheet) + '!A' + FIRST + ':' + L(CFG.A.notes) + MAXROW, G.q(CFG.B.sheet) + '!A' + FIRST + ':' + L(CFG.B.notes) + MAXROW,
  ], 'serial');
  return { A: fresh('A', parseRows('A', dA), normMonth(((mA[0] || [])[0]))), B: fresh('B', parseRows('B', dB), normMonth(((mB[0] || [])[0]))), serverTime: Date.now() };
}

// A row is still the same customer the person was looking at? (guards against two people editing at once)
async function checkRow(cls, row, expectName) {
  const c = cfgOf(cls);
  if (!Number.isInteger(row) || row < FIRST || row > MAXROW) throw new Error('Invalid row');
  const [r] = await G.batchGet(SHEET_ID(), [cell(c, c.name, row)]);
  const name = str(((r[0] || [])[0])).trim();
  if (!name) throw new Error('That customer is no longer in the list — refresh and try again');
  if (expectName !== undefined && str(expectName).trim() !== name) throw new Error('This row changed since you opened it — refresh and try again');
  return name;
}

async function toggle(a) {
  const c = cfgOf(a.cls), t = Number(a.tp);
  if (!(t >= 0 && t < TP) || !Number.isInteger(t)) throw new Error('Invalid touch point');
  await checkRow(a.cls, Number(a.row), a.expect);
  await G.batchUpdate(SHEET_ID(), [{ range: cell(c, c.tp + t, Number(a.row)), values: [[a.done ? 'Done' : '']] }]);
  return true;
}

let propsCache = null;
async function sheetId(title) {
  if (!propsCache) propsCache = await G.sheetProps(SHEET_ID());
  const p = propsCache.find((x) => x.title === title); if (!p) throw new Error('Tab "' + title + '" not found');
  return p.sheetId;
}
async function numberFormat(c, row, col, pattern, type) {
  await G.format(SHEET_ID(), [{ repeatCell: { range: { sheetId: await sheetId(c.sheet), startRowIndex: row - 1, endRowIndex: row, startColumnIndex: col - 1, endColumnIndex: col },
    cell: { userEnteredFormat: { numberFormat: { type: type || 'NUMBER', pattern } } }, fields: 'userEnteredFormat.numberFormat' } }]);
}
const RS = '"Rs. "#,##0';

function clean(a) {
  const name = clip(a.name, 100); if (!name) throw new Error('Customer name is required.');
  return { name, phone: clip(a.phone, 30), state: clip(a.state, 40), value: num(a.value), notes: str(a.notes).trim().slice(0, 600) };
}

async function addClient(a) {
  const c = cfgOf(a.cls), d = clean(a);
  const [rows] = await G.batchGet(SHEET_ID(), [G.q(c.sheet) + '!A' + FIRST + ':' + L(c.name) + MAXROW]);
  // first row with no customer name (the sheet comes with numbered, formula-filled blank rows); skip legend/text rows in column A
  let target = -1;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || [], nm = str(r[c.name - 1]).trim(), n = str(r[c.num - 1]).trim();
    if (!nm && (n === '' || /^\d+$/.test(n))) { target = FIRST + i; break; }
  }
  if (target === -1) target = FIRST + rows.length;
  const valueCol = c.split ? c.avg : c.order;
  await G.batchUpdate(SHEET_ID(), [
    { range: G.q(c.sheet) + '!' + L(c.num) + target + ':' + L(c.state) + target, values: [[target - FIRST + 1, d.name, d.phone, d.state]] },
    { range: cell(c, valueCol, target), values: [[d.value]] },
    { range: cell(c, c.notes, target), values: [[d.notes]] },
    { range: G.q(c.sheet) + '!' + L(c.tp) + target + ':' + L(c.tp + TP - 1) + target, values: [['', '', '', '', '']] },
  ]);
  // completion % and status formulas (the same ones the sheet's own rows use)
  const tpA1 = L(c.tp) + target + ':' + L(c.tp + TP - 1) + target, cnt = 'COUNTIF(' + tpA1 + ',"Done")';
  await G.batchUpdate(SHEET_ID(), [
    { range: cell(c, c.comp, target), values: [['=' + cnt + '/' + TP]] },
    { range: cell(c, c.status, target), values: [['=IF(' + cnt + '=0,"Not Started",IF(' + cnt + '>=2,"Completed","On Track"))']] },
  ], 'USER_ENTERED');
  await numberFormat(c, target, valueCol, RS);
  await numberFormat(c, target, c.comp, '0%', 'PERCENT');
  return { row: target };
}

async function updateClient(a) {
  const c = cfgOf(a.cls), d = clean(a), row = Number(a.row);
  await checkRow(a.cls, row, a.expect);
  const valueCol = c.split ? c.avg : c.order;
  await G.batchUpdate(SHEET_ID(), [
    { range: G.q(c.sheet) + '!' + L(c.name) + row + ':' + L(c.state) + row, values: [[d.name, d.phone, d.state]] },
    { range: cell(c, valueCol, row), values: [[d.value]] },
    { range: cell(c, c.notes, row), values: [[d.notes]] },
  ]);
  await numberFormat(c, row, valueCol, RS);
  return true;
}

const istDate = () => new Intl.DateTimeFormat('en-GB', { timeZone: 'Asia/Kolkata', day: '2-digit', month: 'short', year: 'numeric' }).format(new Date());

// Retire a customer who went quiet or left: log to "Archived Customers", then free the row for a new client.
async function archiveClient(a, who) {
  const c = cfgOf(a.cls), row = Number(a.row);
  await checkRow(a.cls, row, a.expect);
  const [shown] = await G.batchGet(SHEET_ID(), [G.q(c.sheet) + '!A' + row + ':' + L(c.notes) + row], undefined, true);
  const r = (shown[0] || []);
  const valueDisplay = str(r[(c.split ? c.avg : c.order) - 1]);
  const [col] = await G.batchGet(SHEET_ID(), [G.q(ARCHIVED) + '!A1:A' + MAXROW]);
  const insertAt = col.length < 1 ? 2 : col.length + 1;
  await G.batchUpdate(SHEET_ID(), [{ range: G.q(ARCHIVED) + '!A' + insertAt + ':H' + insertAt,
    values: [['Class ' + a.cls, str(r[c.name - 1]), str(r[c.phone - 1]), str(r[c.state - 1]), valueDisplay, clip(a.reason, 200), istDate(), str(r[c.notes - 1])]] }]);
  const blanks = [];
  const clear = (col1, n) => blanks.push({ range: G.q(c.sheet) + '!' + L(col1) + row + (n > 1 ? ':' + L(col1 + n - 1) + row : ''), values: [new Array(n).fill('')] });
  clear(c.name, 3);                                                // name, phone, state
  if (c.split) clear(c.total, 2); else clear(c.order, 1);          // value columns
  clear(c.tp, TP); clear(c.notes, 1);
  await G.batchUpdate(SHEET_ID(), blanks);
  return true;
}

/* ---------- monthly rollover (off unless TOUCH_ROLLOVER=1) ---------- */
async function archiveClassMonth(cls, label) {
  const c = CFG[cls];
  const [shown] = await G.batchGet(SHEET_ID(), [G.q(c.sheet) + '!A' + FIRST + ':' + L(c.notes) + MAXROW], undefined, true);
  const rows = [];
  (shown || []).forEach((r) => {
    if (!str(r[c.name - 1]).trim()) return;
    const tps = []; for (let t = 0; t < TP; t++) tps.push(str(r[c.tp - 1 + t]));
    const done = tps.filter((x) => x.trim() === 'Done').length;
    rows.push([str(r[c.num - 1]), str(r[c.name - 1]), str(r[c.phone - 1]), str(r[c.state - 1]), str(r[(c.split ? c.avg : c.order) - 1]),
      ...tps, str(r[c.comp - 1]) || Math.round((done / TP) * 100) + '%', str(r[c.status - 1]) || statusOf(done), str(r[c.notes - 1])]);
  });
  if (!rows.length) return;
  const [arc] = await G.batchGet(SHEET_ID(), [G.q(ARCHIVE) + '!A' + FIRST + ':B' + MAXROW], 'serial');
  const already = (arc || []).some((r) => normMonth(r[0]) === label && str(r[1]).trim() === 'Class ' + cls);
  if (already) return;                                             // another system already archived this month
  let last = -1; (arc || []).forEach((r, i) => { if (str(r[0]).trim() !== '') last = i; });
  const insertAt = FIRST + last + 1;
  await G.batchUpdate(SHEET_ID(), [{ range: G.q(ARCHIVE) + '!A' + insertAt + ':O' + (insertAt + rows.length - 1), values: rows.map((r) => [label, 'Class ' + cls, ...r]) }]);
}

async function resetTouchpoints(cls) {
  const c = CFG[cls];
  const [names] = await G.batchGet(SHEET_ID(), [G.q(c.sheet) + '!A' + FIRST + ':' + L(c.name) + MAXROW]);
  let lastNamed = -1; (names || []).forEach((r, i) => { if (str((r || [])[c.name - 1]).trim()) lastNamed = i; });
  if (lastNamed < 0) return;
  await G.batchUpdate(SHEET_ID(), [{ range: G.q(c.sheet) + '!' + L(c.tp) + FIRST + ':' + L(c.tp + TP - 1) + (FIRST + lastNamed),
    values: Array.from({ length: lastNamed + 1 }, () => new Array(TP).fill('')) }]);
}

async function rollover() {
  if (process.env.TOUCH_ROLLOVER !== '1') return;
  const now = monthLabel(new Date(), process.env.TOUCH_TZ || 'Asia/Kolkata');
  const [mA, mB] = await G.batchGet(SHEET_ID(), [G.q(CFG.A.sheet) + '!B2', G.q(CFG.B.sheet) + '!B2'], 'serial');
  for (const [cls, m] of [['A', mA], ['B', mB]]) {
    const cur = normMonth((m[0] || [])[0]);
    if (cur === now) continue;
    if (cur) { await archiveClassMonth(cls, cur); await resetTouchpoints(cls); }
    await G.batchUpdate(SHEET_ID(), [{ range: G.q(CFG[cls].sheet) + '!B2', values: [[now]] }]);
  }
}

const ROUTES = {
  'touch.getAll': () => getAll(),
  'touch.toggle': (a) => toggle(a),
  'touch.addClient': (a) => addClient(a),
  'touch.updateClient': (a) => updateClient(a),
  'touch.archiveClient': (a, who) => archiveClient(a, who),
};
module.exports = { ROUTES, SHEET_ID, rollover, _t: { normMonth, monthLabel, parseRows } };
