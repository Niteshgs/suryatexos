// Payroll System: one profile tab per person in the Payroll sheet (ANIL … YOGENDRA, KAILASH).
// Every profile has the same layout; cells are found by their labels, so a small layout drift in one tab does no harm.
//   COMMITMENT   : header row "MONTHS" with year columns, then JAN..DEC rows (1 / 0.5 / 0, blank = not filled)
//   Commitment Yearly / FINAL Commitment Score : formulas in the sheet (never written by the app)
//   PAY SCALE    : "Basic Salary: N" and "Basic for OT: N" (text in column B)
//   BANK DETAILS : labels in column E, values in column F (Notes may use the two rows below)
//   Increment Trend : rows below the header in B (month) and C (amount)
//   PERFORMANCE  : free-text cells directly under Achievements / Goals / Attitude / Skill / Knowledge / Ideas
// Env (optional): PAYROLL_SHEET_ID
const G = require('./google');

const SHEET_ID = () => process.env.PAYROLL_SHEET_ID || '1I0her16mt0w0EQgTEX5TOCHa35qek15r-89tgohFA4Q';
const PEOPLE = ['ANIL', 'AJAY', 'BABLU', 'BIKAS ONE', 'DILKHUSH', 'FANTUSH', 'HANSRAJ', 'KINKAR', 'KUNDAN', 'NILESH', 'PARMESHWAR', 'SANGAM', 'SHAILESH', 'SHASHI', 'UMESH', 'UPENDAR', 'YOGENDRA', 'KAILASH'];
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const MONFULL = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
const str = (v) => (v === null || v === undefined ? '' : String(v));
const norm = (s) => str(s).replace(/\s+/g, ' ').trim();
const colL = (i) => String.fromCharCode(65 + i);
const serialOf = (y, m) => Math.round((Date.UTC(y, m - 1, 1) - Date.UTC(1899, 11, 30)) / 86400000);

let tabCache = { at: 0, map: null, props: null };
async function tabs() {
  if (tabCache.map && Date.now() - tabCache.at < 60000) return tabCache;
  const props = await G.sheetProps(SHEET_ID()), map = {};
  PEOPLE.forEach((p) => { const t = props.find((x) => norm(x.title).toUpperCase() === p); if (t) map[p] = t; });
  tabCache = { at: Date.now(), map, props };
  return tabCache;
}
async function resolve(key) {
  const k = norm(key).toUpperCase(), c = await tabs();
  if (!c.map[k]) throw new Error('Profile not found');
  return c.map[k];
}

// "Mar 2025" / "February 2025" / serial number -> { label, ym }
function whenOf(v) {
  if (typeof v === 'number' && v > 20000) {
    const t = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return { label: MON[t.getUTCMonth()] + ' ' + t.getUTCFullYear(), ym: t.getUTCFullYear() + '-' + String(t.getUTCMonth() + 1).padStart(2, '0') };
  }
  const s = norm(v); if (!s) return null;
  const x = /^([A-Za-z]+)\.?\s*,?\s*(\d{4})$/.exec(s);
  if (x) { const i = MONFULL.findIndex((m) => m.slice(0, 3) === x[1].toLowerCase().slice(0, 3)); if (i >= 0) return { label: MON[i] + ' ' + x[2], ym: x[2] + '-' + String(i + 1).padStart(2, '0') }; }
  return { label: s, ym: '' };
}
const money = (v) => { if (typeof v === 'number') return v; const n = Number(str(v).replace(/[^\d.-]/g, '')); return str(v).trim() && Number.isFinite(n) ? n : null; };

// rows = values of A1:O90 (unformatted). Returns the whole profile and where each part lives.
function parse(rows) {
  const at = (r, c) => (rows[r] && rows[r][c] !== undefined ? rows[r][c] : '');
  const find = (col, re, from) => { for (let r = from || 0; r < rows.length; r++) if (re.test(norm(at(r, col)))) return r; return -1; };
  const out = { name: '', years: [], yearly: {}, final: null, basic: null, ot: null, bank: {}, increments: [], perf: [], at: {} };
  const nm = /^name:\s*(.*)$/i.exec(norm(at(3, 1))); out.name = nm ? nm[1] : norm(at(3, 1));

  const mr = find(1, /^MONTHS$/i);
  if (mr >= 0) {
    for (let c = 2; c < 15; c++) { const y = at(mr, c); if (typeof y === 'number' && y > 2000 && y < 2100) out.years.push({ year: y, col: c, months: [] }); else if (c > 2 && !str(y).trim()) break; }
    out.years.forEach((y) => {
      for (let m = 0; m < 12; m++) { const v = at(mr + 1 + m, y.col); y.months.push(typeof v === 'number' ? v : null); }
      const yr = find(1, /^Commitment Yearly/i, mr); const yv = yr >= 0 ? at(yr, y.col) : ''; out.yearly[y.year] = typeof yv === 'number' ? yv : null;
    });
    out.at.monthsRow = mr;
  }
  const fr = find(1, /^FINAL Commitment/i); if (fr >= 0 && typeof at(fr, 2) === 'number') out.final = at(fr, 2);

  const pr = find(1, /^PAY SCALE/i);
  if (pr >= 0) for (let r = pr + 1; r < pr + 6; r++) {
    const s = norm(at(r, 1)); let x;
    if ((x = /^Basic Salary:\s*(.*)$/i.exec(s))) { out.basic = money(x[1]); out.at.basicRow = r; }
    else if ((x = /^Basic for OT:\s*(.*)$/i.exec(s))) { out.ot = money(x[1]); out.at.otRow = r; }
  }
  const br = find(4, /^BANK DETAILS/i);
  if (br >= 0) {
    const lab = { holder: /^Account Holder/i, bank: /^Bank Name/i, acct: /^Account No/i, ifsc: /^IFSC/i };
    Object.keys(lab).forEach((k) => { const r = find(4, lab[k], br); if (r >= 0) { out.bank[k] = str(at(r, 5)).trim(); out.at[k] = r; } });
    const nr = find(4, /^Notes/i, br);
    if (nr >= 0) { out.at.notes = nr; out.bank.notes = [0, 1, 2].map((i) => str(at(nr + i, 5)).trim()).filter(Boolean); }
  }
  const ir = find(1, /^Increment Trend/i);
  if (ir >= 0) {
    out.at.incRow = ir;
    for (let r = ir + 1; r < Math.min(rows.length, ir + 40); r++) {
      const w = whenOf(at(r, 1)), a = money(at(r, 2)); if (!w && a === null) continue;
      out.increments.push({ label: w ? w.label : '', ym: w ? w.ym : '', amount: a });
    }
  }
  const PL = [['achievements', /^Achievements/i, 'Achievements'], ['goals', /^Goals/i, 'Goals'], ['attitude', /^Attitude/i, 'Attitude'], ['skill', /^Skill/i, 'Skill'], ['knowledge', /^Knowl/i, 'Knowledge'], ['ideas', /^Ideas/i, 'Ideas']];
  PL.forEach((p) => {
    for (let r = 0; r < rows.length; r++) for (let c = 6; c < 15; c++) if (p[1].test(norm(at(r, c)))) { out.perf.push({ key: p[0], label: p[2], row: r, col: c, text: str(at(r + 1, c)) }); r = rows.length; break; }
  });
  return out;
}
// Light-blue fill on a month cell = "increment received this month" (set by hand in the sheet, or from the app).
const MARK = { red: 0.6235294, green: 0.77254903, blue: 0.9098039 }, WHITE = { red: 1, green: 1, blue: 1 };
const isBlue = (c) => { if (!c) return false; const r = c.red || 0, b = c.blue || 0; return b - r >= 0.1 && r < 0.95; };
async function marksGrid(tab, p) {
  if (p.at.monthsRow === undefined) return;
  const r0 = p.at.monthsRow + 2, rng = G.q(tab.title) + '!C' + r0 + ':O' + (r0 + 11);
  const url = 'https://sheets.googleapis.com/v4/spreadsheets/' + encodeURIComponent(SHEET_ID()) + '?ranges=' + encodeURIComponent(rng) + '&fields=' + encodeURIComponent('sheets.data.rowData.values.userEnteredFormat.backgroundColor');
  const r = await fetch(url, { headers: { Authorization: 'Bearer ' + (await G.token()) } });
  if (!r.ok) return;
  const rows = ((((await r.json()).sheets || [])[0] || {}).data || [])[0];
  const rd = (rows && rows.rowData) || [];
  p.years.forEach((y) => { y.marks = []; for (let m = 0; m < 12; m++) { const v = ((rd[m] || {}).values || [])[y.col - 2]; y.marks.push(isBlue(v && v.userEnteredFormat && v.userEnteredFormat.backgroundColor)); } });
}
const read = async (tab) => {
  const p = parse((await G.batchGet(SHEET_ID(), [G.q(tab.title) + '!A1:O90'], 'serial'))[0] || []);
  try { await marksGrid(tab, p); } catch (e) { console.error('payroll marks failed: ' + e.message); }
  return p;
};
const fillReq = (tab, p, year, month, on) => {
  const y = p.years.find((x) => x.year === Number(year)); if (!y) return null;
  const row = p.at.monthsRow + month;   // 0-based: months row + month (JAN = +1)
  return { repeatCell: { range: { sheetId: tab.sheetId, startRowIndex: row, endRowIndex: row + 1, startColumnIndex: y.col, endColumnIndex: y.col + 1 },
    cell: { userEnteredFormat: { backgroundColor: on ? MARK : WHITE } }, fields: 'userEnteredFormat.backgroundColor' } };
};

function curYearIdx(p) {
  const nowY = new Date(Date.now() + 5.5 * 3600e3).getUTCFullYear();
  const i = p.years.findIndex((y) => y.year === nowY); return i >= 0 ? i : p.years.length - 1;
}
function bankFlags(p) { return ['holder', 'bank', 'acct', 'ifsc'].filter((k) => !str(p.bank[k]).trim()); }

async function list() {
  const c = await tabs(), keys = PEOPLE.filter((k) => c.map[k]);
  const sets = await G.batchGet(SHEET_ID(), keys.map((k) => G.q(c.map[k].title) + '!A1:O90'), 'serial');
  const now = new Date(Date.now() + 5.5 * 3600e3);
  const people = keys.map((k, i) => {
    const p = parse(sets[i] || []), ci = curYearIdx(p), cy = p.years[ci];
    return { key: k, name: p.name || k, final: p.final, year: cy ? cy.year : null, yearScore: cy ? p.yearly[cy.year] : null,
      basic: p.basic, missingBank: bankFlags(p).length, increments: p.increments.length };
  });
  return { people, year: now.getUTCFullYear() };
}

async function get(a) {
  const tab = await resolve(a.key), p = await read(tab);
  return { key: norm(a.key).toUpperCase(), name: p.name, years: p.years, yearly: p.yearly, final: p.final, basic: p.basic, ot: p.ot,
    bank: { holder: p.bank.holder || '', bank: p.bank.bank || '', acct: p.bank.acct || '', ifsc: p.bank.ifsc || '', notes: [0, 1, 2].map((i) => (p.bank.notes || [])[i] || '') },
    increments: p.increments, perf: p.perf.map((x) => ({ key: x.key, label: x.label, text: x.text })), currentYear: p.years[curYearIdx(p)] ? p.years[curYearIdx(p)].year : null };
}

const clean = (s, n) => str(s).replace(/\r/g, '').replace(/[ \t]+/g, ' ').trim().slice(0, n || 120);
const upd = (tab, cell, v) => ({ range: G.q(tab.title) + '!' + cell, values: [[v]] });

async function setCommit(a) {
  const tab = await resolve(a.key), p = await read(tab);
  const yr = p.years.find((y) => y.year === Number(a.year)); if (!yr) throw new Error('That year is not in this profile');
  const m = Number(a.month); if (!(m >= 1 && m <= 12)) throw new Error('Bad month');
  let v = a.value; if (v === null || v === '' || v === undefined) v = ''; else { v = Number(v); if (![0, 0.5, 1].includes(v)) throw new Error('Commitment must be 1, 0.5 or 0'); }
  await G.batchUpdate(SHEET_ID(), [upd(tab, colL(yr.col) + (p.at.monthsRow + 1 + m), v)], 'RAW');
  return get(a);
}
async function setMark(a) {
  const tab = await resolve(a.key), p = await read(tab), m = Number(a.month);
  if (!(m >= 1 && m <= 12)) throw new Error('Bad month');
  const rq = fillReq(tab, p, a.year, m, !!a.on); if (!rq) throw new Error('That year is not in this profile');
  await G.format(SHEET_ID(), [rq]);
  return get(a);
}
async function setPay(a) {
  const tab = await resolve(a.key), p = await read(tab), d = [];
  const num = (x) => { if (x === '' || x === null || x === undefined) return ''; const n = Number(x); if (!Number.isFinite(n) || n < 0 || n > 1e7) throw new Error('Please enter a valid amount'); return Math.round(n * 100) / 100; };
  if (p.at.basicRow === undefined || p.at.otRow === undefined) throw new Error('Pay scale rows were not found in this profile');
  const b = num(a.basic), o = num(a.ot);
  d.push(upd(tab, 'B' + (p.at.basicRow + 1), 'Basic Salary: ' + b), upd(tab, 'B' + (p.at.otRow + 1), 'Basic for OT: ' + o));
  await G.batchUpdate(SHEET_ID(), d, 'RAW');
  return get(a);
}
async function setBank(a) {
  const tab = await resolve(a.key), p = await read(tab), d = [], b = a.bank || {};
  ['holder', 'bank', 'acct', 'ifsc'].forEach((k) => {
    if (p.at[k] === undefined) throw new Error('Bank rows were not found in this profile');
    let v = clean(b[k], 60); if (k === 'ifsc') v = v.toUpperCase().replace(/\s/g, ''); if (k === 'acct') v = v.replace(/\s/g, '');
    d.push(upd(tab, 'F' + (p.at[k] + 1), v));
  });
  if (p.at.notes === undefined) throw new Error('Bank rows were not found in this profile');
  const n = Array.isArray(b.notes) ? b.notes : [];
  [0, 1, 2].forEach((i) => d.push(upd(tab, 'F' + (p.at.notes + 1 + i), clean(n[i], 120))));
  await G.batchUpdate(SHEET_ID(), d, 'RAW');
  return get(a);
}
async function setIncrements(a) {
  const tab = await resolve(a.key), p = await read(tab);
  if (p.at.incRow === undefined) throw new Error('Increment Trend was not found in this profile');
  const items = (Array.isArray(a.items) ? a.items : []).slice(0, 30).map((x) => {
    const m = /^(\d{4})-(\d{2})$/.exec(str(x.ym)); const amt = Number(x.amount);
    if (!m || +m[2] < 1 || +m[2] > 12 || +m[1] < 2000 || +m[1] > 2100) throw new Error('Please pick a valid month for each increment');
    if (!Number.isFinite(amt) || amt <= 0 || amt > 1e7) throw new Error('Please enter a valid amount for each increment');
    return { y: +m[1], m: +m[2], amt: Math.round(amt * 100) / 100 };
  }).sort((x, y) => x.y - y.y || x.m - y.m);
  const first = p.at.incRow + 2, last = p.at.incRow + 41;
  await G.clear(SHEET_ID(), [G.q(tab.title) + '!B' + first + ':C' + last]);
  if (items.length) {
    await G.batchUpdate(SHEET_ID(), [{ range: G.q(tab.title) + '!B' + first + ':C' + (first + items.length - 1), values: items.map((x) => [serialOf(x.y, x.m), x.amt]) }], 'RAW');
    const rng = (c0, c1) => ({ sheetId: tab.sheetId, startRowIndex: first - 1, endRowIndex: first - 1 + items.length, startColumnIndex: c0, endColumnIndex: c1 });
    await G.format(SHEET_ID(), [
      { repeatCell: { range: rng(1, 3), cell: { userEnteredFormat: { textFormat: { fontFamily: 'Roboto', fontSize: 11 }, horizontalAlignment: 'LEFT' } }, fields: 'userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment' } },
      { repeatCell: { range: rng(1, 2), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } },
    ]);
  }
  if (a.mark && items.length) {
    const rq = items.map((x) => fillReq(tab, p, x.y, x.m, true)).filter(Boolean);
    if (rq.length) await G.format(SHEET_ID(), rq);
  }
  if (a.setBasic && items.length && p.at.basicRow !== undefined) {
    await G.batchUpdate(SHEET_ID(), [upd(tab, 'B' + (p.at.basicRow + 1), 'Basic Salary: ' + items[items.length - 1].amt)], 'RAW');
  }
  return get(a);
}
async function setPerf(a) {
  const tab = await resolve(a.key), p = await read(tab), d = [];
  const t = a.perf && typeof a.perf === 'object' ? a.perf : {};
  p.perf.forEach((x) => { if (Object.prototype.hasOwnProperty.call(t, x.key)) d.push(upd(tab, colL(x.col) + (x.row + 2), clean(t[x.key], 500))); });
  await G.batchUpdate(SHEET_ID(), d, 'RAW');
  return get(a);
}

const ROUTES = {
  'payroll.list': () => list(), 'payroll.get': (a) => get(a || {}),
  'payroll.setCommit': (a) => setCommit(a || {}), 'payroll.setMark': (a) => setMark(a || {}), 'payroll.setPay': (a) => setPay(a || {}), 'payroll.setBank': (a) => setBank(a || {}),
  'payroll.setIncrements': (a) => setIncrements(a || {}), 'payroll.setPerf': (a) => setPerf(a || {}),
};
module.exports = { ROUTES, _t: { parse, whenOf } };
