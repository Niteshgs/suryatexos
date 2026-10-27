// Cash Expenses — direct Google Sheets access, same sheet and layout the team's Apps Script app uses.
// Sheet1: header row found by scanning column A for "Date" (fallback row 5); columns A Date, B Month (1st of month), C Description,
// D Category, E Amount, F Notes. Column G ("Total Monthly") holds the owner's own formulas and is never written.
// Env: EXPENSES_SHEET_ID (default = "Cash Expenses Suryatex V2"), EXPENSES_TAB (default Sheet1)
const G = require('./google');

const SHEET_ID = () => process.env.EXPENSES_SHEET_ID || '16-CAOYnjoBx7YWtI1xjqgjZIWva0t-izkKx3jXXPAbo';
const TAB = () => process.env.EXPENSES_TAB || 'Sheet1';
const CATEGORIES = ['Office Supplies', 'Freight', 'Meals / Pooja', 'Utilities/Gas', 'Spares', 'Hamali', 'Miscellaneous', 'Fuel'];
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const MAXROW = 5000;
const str = (v) => (v === null || v === undefined ? '' : String(v));

const serialOf = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
function ymd(v) {                                   // serial number, 'yyyy-mm-dd' or 'm/d/yyyy' -> {y,m,d} or null
  if (typeof v === 'number' && v > 0) { const t = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000); return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() }; }
  const s = str(v).trim(); let x;
  if ((x = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s))) return { y: +x[1], m: +x[2], d: +x[3] };
  if ((x = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s))) return { y: +x[3], m: +x[1], d: +x[2] };
  return null;
}
const fmtDay = (v) => { const t = ymd(v); return t ? String(t.d).padStart(2, '0') + ' ' + MON[t.m - 1] + ' ' + t.y : str(v); };
const fmtMonth = (v) => { const t = ymd(v); return t ? MON[t.m - 1] + ' ' + t.y : str(v); };

async function headerRow() {
  const col = (await G.batchGet(SHEET_ID(), [G.q(TAB()) + '!A1:A60']))[0] || [];
  for (let i = 0; i < col.length; i++) if (str((col[i] || [])[0]).trim().toLowerCase() === 'date') return i + 1;
  return 5;
}
async function readRows() {
  const h = await headerRow();
  const rows = (await G.batchGet(SHEET_ID(), [G.q(TAB()) + '!A' + (h + 1) + ':F' + MAXROW], 'serial'))[0] || [];
  return { h, rows };
}

async function add(a) {
  const t = ymd(a.date); if (!t) throw new Error('Please pick a valid date');
  const desc = str(a.description).trim().slice(0, 200); if (!desc) throw new Error('Please enter a description');
  const cat = str(a.category).trim().slice(0, 60); if (!cat) throw new Error('Please select a category');
  const amount = Number(a.amount); if (!Number.isFinite(amount) || amount <= 0 || amount > 1e8) throw new Error('Please enter a valid amount');
  const notes = str(a.notes).trim().slice(0, 500);
  const { h, rows } = await readRows();
  let last = 0; rows.forEach((r, i) => { if (r.some((c) => str(c).trim() !== '')) last = i + 1; });
  const row = h + last + 1;
  const tab = G.q(TAB());
  // dates as real date serials (what the old script's appendRow produced); text cells are written raw so nothing is mangled
  await G.batchUpdate(SHEET_ID(), [{ range: tab + '!A' + row + ':F' + row, values: [[serialOf(t.y, t.m, t.d), serialOf(t.y, t.m, 1), desc, cat, Math.round(amount * 100) / 100, notes]] }], 'RAW');
  const sid = (await G.sheetProps(SHEET_ID())).filter((p) => p.title === TAB())[0];
  if (sid) {
    const fm = (c0, c1, type, pattern) => ({ repeatCell: { range: { sheetId: sid.sheetId, startRowIndex: row - 1, endRowIndex: row, startColumnIndex: c0, endColumnIndex: c1 },
      cell: { userEnteredFormat: { numberFormat: { type, pattern } } }, fields: 'userEnteredFormat.numberFormat' } });
    await G.format(SHEET_ID(), [fm(0, 2, 'DATE', 'M/d/yyyy'), fm(4, 5, 'CURRENCY', '[$₹]#,##0.00')]);
  }
  return { row };
}
async function list() {
  const { rows } = await readRows();
  const out = [];
  rows.forEach((r) => {
    if (!r[0] && !r[2] && !r[4]) return;
    const amt = typeof r[4] === 'number' ? r[4] : Number(str(r[4]).replace(/[^\d.-]/g, '')) || 0;
    out.push({ ts: (ymd(r[0]) ? serialOf(ymd(r[0]).y, ymd(r[0]).m, ymd(r[0]).d) : 0), date: fmtDay(r[0]), description: str(r[2]), category: str(r[3]), amount: amt, notes: str(r[5]) });
  });
  out.sort((a, b) => b.ts - a.ts);
  return { rows: out.map(({ ts, ...x }) => x), categories: CATEGORIES };
}
async function totals() {
  const { rows } = await readRows();
  const m = {}, order = [];
  rows.forEach((r) => {
    if (!r[1] && !r[4]) return;
    const key = fmtMonth(r[1]); const amt = typeof r[4] === 'number' ? r[4] : Number(str(r[4]).replace(/[^\d.-]/g, '')) || 0;
    if (!(key in m)) { m[key] = 0; order.push(key); }
    m[key] += amt;
  });
  return { totals: order.map((k) => ({ month: k, total: Math.round(m[k] * 100) / 100 })) };
}
const ROUTES = {
  'expenses.add': (a) => add(a || {}),
  'expenses.list': () => list(),
  'expenses.totals': () => totals(),
};
module.exports = { ROUTES, CATEGORIES };
