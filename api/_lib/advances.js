// Advances: one row per advance taken, in the ADVANCES tab of the Payroll sheet.
// Columns: A Date | B Name | C Advance Amount | D Entered by | E Entered at. Only this tab is touched.
// Env (optional): ADVANCES_SHEET_ID, ADVANCES_TAB
const G = require('./google');

const SHEET_ID = () => process.env.ADVANCES_SHEET_ID || '1I0her16mt0w0EQgTEX5TOCHa35qek15r-89tgohFA4Q';
const TAB = () => process.env.ADVANCES_TAB || 'ADVANCES';
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const str = (v) => (v === null || v === undefined ? '' : String(v));
const serialOf = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
function ymd(v) {
  if (typeof v === 'number' && v > 0) { const t = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000); return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() }; }
  const s = str(v).trim(); let x;
  if ((x = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s))) return { y: +x[1], m: +x[2], d: +x[3] };
  if ((x = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s))) return { y: +x[3], m: +x[1], d: +x[2] };
  return null;
}
const fmtDay = (t) => String(t.d).padStart(2, '0') + ' ' + MON[t.m - 1] + ' ' + t.y;
const nowIST = () => new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata', day: '2-digit', month: 'short', year: 'numeric', hour: 'numeric', minute: '2-digit', hour12: true });
const clean = (s) => str(s).replace(/\s+/g, ' ').trim().slice(0, 80);   // written RAW, so nothing is read as a formula

async function add(a, who) {
  const t = ymd(a.date); if (!t) throw new Error('Please pick a valid date');
  const name = clean(a.name); if (!name) throw new Error('Please enter the name');
  const amount = Number(a.amount); if (!Number.isFinite(amount) || amount <= 0 || amount > 1e7) throw new Error('Please enter a valid amount');
  const titles = await G.sheetTitles(SHEET_ID());
  if (!titles.includes(TAB())) throw new Error('The ADVANCES tab was not found in the sheet');
  const amt = Math.round(amount * 100) / 100;
  await G.append(SHEET_ID(), G.q(TAB()) + '!A:E', [[serialOf(t.y, t.m, t.d), name, amt, str(who), nowIST()]], 'RAW');
  return { saved: true, name, amount: amt, date: fmtDay(t) };
}

async function list() {
  const rows = (await G.batchGet(SHEET_ID(), [G.q(TAB()) + '!A2:E5000'], 'serial'))[0] || [];
  const out = [], names = new Set(), now = new Date(Date.now() + 5.5 * 3600e3);
  const cy = now.getUTCFullYear(), cm = now.getUTCMonth() + 1;
  let month = 0;
  rows.forEach((r) => {
    const t = ymd(r[0]), name = str(r[1]).trim(), amt = typeof r[2] === 'number' ? r[2] : Number(str(r[2]).replace(/[^\d.-]/g, '')) || 0;
    if (!t && !name && !amt) return;
    if (name) names.add(name);
    if (t && t.y === cy && t.m === cm) month += amt;
    out.push({ date: t ? fmtDay(t) : str(r[0]), name, amount: amt, by: str(r[3]), at: str(r[4]) });
  });
  return { rows: out.reverse().slice(0, 150), names: [...names].sort((x, y) => x.localeCompare(y)), month: Math.round(month * 100) / 100, count: out.length };
}

const ROUTES = { 'advances.add': (a, who) => add(a || {}, who), 'advances.list': () => list() };
module.exports = { ROUTES };
