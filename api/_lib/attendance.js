// Staff Attendance (+ Leave Plan Form) — direct Google Sheets access, same tabs and layout the team's Apps Script app uses.
// Tabs: Workers (Name, Status) · Attendance (Worker, Date 'yyyy-mm-dd' text, Status, OT, Remarks) · Activity Log · Register (current month grid)
//       · Yearly Register (archive of past months) · Leave Plan Form.
// Env: ATT_SHEET_ID (default = the "Attendance Register" workbook)
const G = require('./google');

const SHEET_ID = () => process.env.ATT_SHEET_ID || '1p6PVxPVGNpIMwGKcRitI8966Ui3zRjX3RkuMJgE9kd0';
const T = { workers: 'Workers', att: 'Attendance', log: 'Activity Log', reg: 'Register', yearly: 'Yearly Register', leave: 'Leave Plan Form' };
const MONTHS_FULL = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const TZ = 'Asia/Kolkata';
const str = (v) => (v === null || v === undefined ? '' : String(v));
const q = (n) => G.q(n);
const L = (n) => { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
const pad = (n) => String(n).padStart(2, '0');
const YMD = /^(\d{4})-(\d{2})-(\d{2})$/;
const daysIn = (y, m) => new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
const monthName = (y, m) => MONTHS_FULL[m] + ' ' + y;
const safe = (s) => (/^[=+\-@]/.test(str(s)) ? "'" + str(s) : str(s));        // stops text being read as a formula when typed in
const text = (s) => "'" + str(s);                                              // keeps date-looking text as text

// A date cell is normally text 'yyyy-mm-dd', but Sheets sometimes turns it into a real date (serial number) or 'm/d/yyyy'.
function dateKey(v) {
  if (typeof v === 'number' && v > 0) { const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000); return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
  const s = str(v).trim(); let m;
  if (YMD.test(s)) return s;
  if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s))) return m[3] + '-' + pad(+m[1]) + '-' + pad(+m[2]);
  return s;
}
function normStatus(raw) { const s = str(raw).trim().toLowerCase(); return s === 'present' || s === 'p' ? 'present' : s === 'absent' || s === 'a' ? 'absent' : s; }
function nowIST() {   // wall-clock "now" in India as a spreadsheet date-time string
  const p = new Intl.DateTimeFormat('en-US', { timeZone: TZ, year: 'numeric', month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }).formatToParts(new Date()).reduce((o, x) => (o[x.type] = x.value, o), {});
  return p.month + '/' + p.day + '/' + p.year + ' ' + (p.hour === '24' ? '00' : p.hour) + ':' + p.minute + ':' + p.second;
}
function todayIST() { const p = new Intl.DateTimeFormat('en-CA', { timeZone: TZ }).format(new Date()); return p; }  // yyyy-mm-dd

async function logRow(op, action, worker, date, oldV, newV) {
  try { await ensureTimeFmt(T.log); const ur = await G.append(SHEET_ID(), q(T.log) + '!A1', [[nowIST(), safe(op || 'unidentified'), action, safe(worker), text(date || ''), safe(oldV), safe(newV)]], 'USER_ENTERED'); await fmtAppended(T.log, ur); } catch (e) { /* logging must never block a save */ }
}

async function readAll() {
  const [w, a] = await G.batchGet(SHEET_ID(), [q(T.workers) + '!A2:B2000', q(T.att) + '!A2:E30000'], 'serial');
  return { w, a };
}
function parse({ w, a }) {
  const workers = [], roster = [];
  w.forEach((r) => { const name = str(r[0]).trim(); if (!name) return; const status = str(r[1]).trim() || 'Active'; roster.push({ name, status }); if (status !== 'Archived') workers.push(name); });
  const attendance = {};
  a.forEach((r) => { if (!r[0] || !r[1]) return; attendance[str(r[0]) + '|' + dateKey(r[1])] = { status: normStatus(r[2]), ot: Number(r[3]) || 0, remarks: str(r[4]) }; });
  return { workers, roster, attendance };
}
async function getAll() { return parse(await readAll()); }

/* ---------- Register (current month grid) ---------- */
const code = (status, ot) => (status === 'present' ? Number(ot) || 0 : status === 'absent' ? 'A' : '');

async function snapshotToYearly(sid, titleRowsCount) {
  try {
    const g = (await G.batchGet(SHEET_ID(), [q(T.reg) + '!A1:AF1000']))[0] || [];
    let n = g.length; while (n > 0 && !g[n - 1].some((c) => str(c) !== '')) n--;
    if (!n) return;
    const colA = (await G.batchGet(SHEET_ID(), [q(T.yearly) + '!A1:A5000']))[0] || [];
    let last = colA.length; while (last > 0 && str((colA[last - 1] || [])[0]) === '') last--;
    const start = last + 1 + (last + 1 > 1 ? 2 : 0);
    const width = Math.max(...g.slice(0, n).map((r) => r.length), 1);
    const rows = g.slice(0, n).map((r) => { const o = r.slice(); while (o.length < width) o.push(''); return o; });
    await G.batchUpdate(SHEET_ID(), [{ range: q(T.yearly) + '!A' + start + ':' + L(width) + (start + n - 1), values: rows }], 'RAW');
    if (sid != null) await G.format(SHEET_ID(), [
      { repeatCell: { range: { sheetId: sid, startRowIndex: start - 1, endRowIndex: start, startColumnIndex: 0, endColumnIndex: 1 }, cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 12 } } }, fields: 'userEnteredFormat.textFormat' } },
      { repeatCell: { range: { sheetId: sid, startRowIndex: start, endRowIndex: start + 1, startColumnIndex: 0, endColumnIndex: width }, cell: { userEnteredFormat: { textFormat: { bold: true } } }, fields: 'userEnteredFormat.textFormat' } },
    ]);
  } catch (e) { /* archiving must never block the Register rebuild */ }
}
const timeFmtDone = {};
// Timestamps are written as real date-times; make sure the column shows them as dates (not serial numbers like 46322.8113).
async function ensureTimeFmt(tab) {
  if (timeFmtDone[tab]) return;
  timeFmtDone[tab] = true;
  try {
    const sid = (await sheetIds())[tab];
    if (sid == null) { timeFmtDone[tab] = false; return; }
    await G.format(SHEET_ID(), [{ repeatCell: { range: { sheetId: sid, startRowIndex: 1, startColumnIndex: 0, endColumnIndex: 1 }, cell: { userEnteredFormat: { numberFormat: { type: 'DATE_TIME', pattern: 'yyyy-mm-dd hh:mm:ss' } } }, fields: 'userEnteredFormat.numberFormat' } }]);
  } catch (e) { timeFmtDone[tab] = false; /* cosmetic only */ }
}
// Format exactly the row that was just appended (newly added grid rows do not inherit the column format).
async function fmtAppended(tab, updatedRange) {
  try {
    const m = /!A(\d+)(?::|$)/.exec(str(updatedRange)); if (!m) return;
    const sid = (await sheetIds())[tab]; if (sid == null) return;
    const r = +m[1] - 1;
    await G.format(SHEET_ID(), [{ repeatCell: { range: { sheetId: sid, startRowIndex: r, endRowIndex: r + 1, startColumnIndex: 0, endColumnIndex: 1 }, cell: { userEnteredFormat: { numberFormat: { type: 'DATE_TIME', pattern: 'yyyy-mm-dd hh:mm:ss' } } }, fields: 'userEnteredFormat.numberFormat' } }]);
  } catch (e) { /* cosmetic only */ }
}
async function sheetIds() { const o = {}; (await G.sheetProps(SHEET_ID())).forEach((p) => { o[p.title] = p.sheetId; }); return o; }

async function rebuildRegister(year, month, op) {
  const ids = await sheetIds();
  if (ids[T.reg] == null) throw new Error('The "Register" tab is missing from the sheet');
  const title = str(((await G.batchGet(SHEET_ID(), [q(T.reg) + '!A1']))[0] || [[]])[0][0]);
  const mn = monthName(year, month);
  if (title && title.indexOf(mn) === -1 && ids[T.yearly] != null) await snapshotToYearly(ids[T.yearly]);
  const { workers, attendance } = parse(await readAll());
  const n = daysIn(year, month);
  const grid = [['ATTENDANCE REGISTER — ' + mn], ['Employee'].concat(Array.from({ length: n }, (_, i) => i + 1))];
  workers.forEach((w) => grid.push([w].concat(Array.from({ length: n }, (_, i) => { const r = attendance[w + '|' + year + '-' + pad(month + 1) + '-' + pad(i + 1)]; return r ? code(r.status, r.ot) : ''; }))));
  await G.clear(SHEET_ID(), [q(T.reg) + '!A1:AF1000']);
  await G.batchUpdate(SHEET_ID(), [{ range: q(T.reg) + '!A1:' + L(n + 1) + grid.length, values: grid }], 'RAW');
  const sid = ids[T.reg];
  await G.format(SHEET_ID(), [
    { repeatCell: { range: { sheetId: sid }, cell: { userEnteredFormat: {} }, fields: 'userEnteredFormat' } },
    { repeatCell: { range: { sheetId: sid, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 1 }, cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 12 } } }, fields: 'userEnteredFormat.textFormat' } },
    { repeatCell: { range: { sheetId: sid, startRowIndex: 1, endRowIndex: 2 }, cell: { userEnteredFormat: { textFormat: { bold: true }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } },
    { repeatCell: { range: { sheetId: sid, startRowIndex: 2, endRowIndex: grid.length, startColumnIndex: 1, endColumnIndex: n + 1 }, cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat.horizontalAlignment' } },
    { repeatCell: { range: { sheetId: sid, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 1 }, cell: { userEnteredFormat: { textFormat: { bold: true } } }, fields: 'userEnteredFormat.textFormat' } },
    { updateSheetProperties: { properties: { sheetId: sid, gridProperties: { frozenRowCount: 2, frozenColumnCount: 1 } }, fields: 'gridProperties(frozenRowCount,frozenColumnCount)' } },
  ]);
  await logRow(op, 'SYNC VIEWS', '', '', '', mn + ' — Register rebuilt (' + workers.length + ' active staff)');
  return { ok: true, month: mn };
}

// Make sure Register shows the month; returns true if it had to rebuild (which already includes every saved mark).
async function ensureRegister(year, month, op) {
  const colA = (await G.batchGet(SHEET_ID(), [q(T.reg) + '!A1:A1000']))[0] || [];
  if (str((colA[0] || [])[0]).indexOf(monthName(year, month)) === -1) { await rebuildRegister(year, month, op); return { rebuilt: true }; }
  return { rebuilt: false, colA };
}
async function registerRow(colA, worker) {
  for (let i = 2; i < colA.length; i++) if (str((colA[i] || [])[0]) === worker) return i + 1;
  const range = await G.append(SHEET_ID(), q(T.reg) + '!A1', [[worker]], 'RAW');
  const m = /!A(\d+)/.exec(range); if (!m) throw new Error('Could not add the Register row');
  return +m[1];
}
async function syncCell(worker, ds, status, ot, op) {
  try {
    const [y, m, d] = ds.split('-').map(Number);
    const r = await ensureRegister(y, m - 1, op); if (r.rebuilt) return;
    const row = await registerRow(r.colA, worker);
    await G.batchUpdate(SHEET_ID(), [{ range: q(T.reg) + '!' + L(d + 1) + row, values: [[code(status, ot)]] }], 'RAW');
  } catch (e) { /* a Register hiccup must never block the real save */ }
}

/* ---------- writes ---------- */
function checkWorker(w) { w = str(w).trim(); if (!w || w.length > 60) throw new Error('Invalid staff name'); return w; }
function checkDate(d) { d = str(d).trim(); const m = YMD.exec(d); if (!m || +m[2] < 1 || +m[2] > 12 || +m[3] < 1 || +m[3] > daysIn(+m[1], +m[2] - 1)) throw new Error('Invalid date'); return d; }
function checkStatus(s) { s = str(s); if (s !== '' && s !== 'present' && s !== 'absent') throw new Error('Invalid status'); return s; }
function checkOT(v, status) { const n = Math.round(Number(v) || 0); if (n < 0 || n > 24) throw new Error('OT must be between 0 and 24 hours'); return status === 'present' ? n : (status === '' ? n : 0); }

let rosterCache = { at: 0, list: [] };
// After a save for today: if that left nobody unmarked, send the one-time "attendance complete" phone alert. Never blocks or fails the save.
async function maybeAlertDone(rows, idx, worker, ds, status, actor) {
  try {
    const today = new Date(Date.now() + 5.5 * 3600 * 1000).toISOString().slice(0, 10);
    if (ds !== today || !status) return;
    const P = require('./push');
    if (!(await P.pending(today))) return;
    if (Date.now() - rosterCache.at > 300000) {
      const w = (await G.batchGet(SHEET_ID(), [q(T.workers) + '!A2:B2000']))[0] || [];
      rosterCache = { at: Date.now(), list: w.filter((r) => str(r[0]).trim() && (str(r[1]).trim() || 'Active') !== 'Archived').map((r) => str(r[0]).trim()) };
    }
    const st = {}; rows.forEach((r) => { if (dateKey(r[1]) === today) st[str(r[0])] = normStatus(r[2]); });
    st[worker] = status;
    const list = rosterCache.list; if (!list.length || !list.includes(worker)) return;
    if (list.some((n) => !st[n])) return;
    await P.attendanceDone(today, { present: list.filter((n) => st[n] === 'present').length, absent: list.filter((n) => st[n] === 'absent').length, total: list.length }, actor);
  } catch (e) { /* an alert problem must never block saving attendance */ }
}

async function save(a, op, actor) {
  const worker = checkWorker(a.worker), ds = checkDate(a.date), status = checkStatus(a.status), ot = checkOT(a.ot, status), remarks = str(a.remarks).trim().slice(0, 200);
  const rows = (await G.batchGet(SHEET_ID(), [q(T.att) + '!A2:E30000'], 'serial'))[0] || [];
  let idx = -1; for (let i = 0; i < rows.length; i++) if (str(rows[i][0]) === worker && dateKey(rows[i][1]) === ds) { idx = i; break; }
  const fmt = (s, o, r) => (s || '-') + ' / OT ' + (o || 0) + (r ? ' / "' + r + '"' : '');
  if (idx >= 0) {
    const r = rows[idx], row = idx + 2;
    await G.batchUpdate(SHEET_ID(), [{ range: q(T.att) + '!C' + row + ':E' + row, values: [[status, ot, remarks]] }], 'RAW');
    await Promise.all([logRow(op, 'EDIT', worker, ds, fmt(str(r[2]), r[3], str(r[4])), fmt(status, ot, remarks)), syncCell(worker, ds, status, ot, op)]);
  } else {
    await G.append(SHEET_ID(), q(T.att) + '!A1', [[worker, ds, status, ot, remarks]], 'RAW');
    await Promise.all([logRow(op, 'MARK', worker, ds, '', fmt(status, ot, remarks)), syncCell(worker, ds, status, ot, op)]);
  }
  await maybeAlertDone(rows, idx, worker, ds, status, actor);
  return { ok: true };
}

async function writeRegisterRange(op, byMonth, worker, valuesFor) {   // byMonth: {key:{year,month,days:[]}}
  try {
    for (const k of Object.keys(byMonth)) {
      const { year, month } = byMonth[k];
      const r = await ensureRegister(year, month, op); if (r.rebuilt) continue;
      const row = await registerRow(r.colA, worker);
      const { from, vals } = valuesFor(byMonth[k]);
      await G.batchUpdate(SHEET_ID(), [{ range: q(T.reg) + '!' + L(from + 1) + row + ':' + L(from + vals.length) + row, values: [vals] }], 'RAW');
    }
  } catch (e) { /* never block the real save */ }
}

async function bulk(a, op) {
  const worker = checkWorker(a.worker), status = checkStatus(a.status), ot = checkOT(a.ot, status), remarks = str(a.remarks).trim().slice(0, 200);
  const s = checkDate(a.start), e = checkDate(a.end);
  if (e < s) throw new Error('Invalid date range');
  const list = []; { const d = new Date(s + 'T00:00:00Z'), end = new Date(e + 'T00:00:00Z'); for (; d <= end; d.setUTCDate(d.getUTCDate() + 1)) list.push(d.toISOString().slice(0, 10)); }
  if (list.length > 62) throw new Error('Range too large — please split into smaller chunks (max ~2 months)');
  const rows = (await G.batchGet(SHEET_ID(), [q(T.att) + '!A2:E30000'], 'serial'))[0] || [];
  const index = {}; rows.forEach((r, i) => { if (str(r[0]) === worker) index[dateKey(r[1])] = i + 2; });
  const upd = [], add = [];
  list.forEach((ds) => { if (index[ds]) upd.push({ range: q(T.att) + '!C' + index[ds] + ':E' + index[ds], values: [[status, ot, remarks]] }); else add.push([worker, ds, status, ot, remarks]); });
  await G.batchUpdate(SHEET_ID(), upd, 'RAW');
  if (add.length) await G.append(SHEET_ID(), q(T.att) + '!A1', add, 'RAW');
  const byMonth = {}; list.forEach((ds) => { const [y, m, d] = ds.split('-').map(Number); (byMonth[y + '-' + m] = byMonth[y + '-' + m] || { year: y, month: m - 1, days: [] }).days.push(d); });
  await writeRegisterRange(op, byMonth, worker, (b) => { const mn = Math.min(...b.days), mx = Math.max(...b.days); return { from: mn, vals: new Array(mx - mn + 1).fill(code(status, ot)) }; });
  await logRow(op, 'BULK MARK', worker, s + ' → ' + e, '', (status || '-') + ' × ' + list.length + ' days');
  return { ok: true, count: list.length };
}

async function monthSave(a, op) {
  const worker = checkWorker(a.worker), year = Math.round(Number(a.year)), month = Math.round(Number(a.month));
  if (!(year >= 2020 && year <= 2100 && month >= 0 && month <= 11)) throw new Error('Invalid month');
  if (!Array.isArray(a.entries) || a.entries.length > 31) throw new Error('Invalid entries');
  const n = daysIn(year, month);
  const rows = (await G.batchGet(SHEET_ID(), [q(T.att) + '!A2:E30000'], 'serial'))[0] || [];
  const index = {}; rows.forEach((r, i) => { if (str(r[0]) === worker) index[dateKey(r[1])] = i + 2; });
  // Conflict check: refuse if someone else changed a day since this person opened the month (unless they chose "save anyway").
  if (a.base && typeof a.base === 'object' && !a.force) {
    const cur = {}; rows.forEach((r) => { if (str(r[0]) === worker) cur[dateKey(r[1])] = { status: normStatus(r[2]), ot: Number(r[3]) || 0, remarks: str(r[4]).trim() }; });
    const same = (x, y) => str(x.status) === str(y.status) && (x.status !== 'present' || (Number(x.ot) || 0) === (Number(y.ot) || 0)) && str(x.remarks).trim() === str(y.remarks).trim();
    const bad = [];
    a.entries.forEach((e) => {
      const day = Math.round(Number(e.day)); if (!(day >= 1 && day <= n) || !a.base[day]) return;
      const c = cur[year + '-' + pad(month + 1) + '-' + pad(day)] || { status: '', ot: 0, remarks: '' };
      const st = normStatus(e.status || ''), mine = { status: st, ot: st === 'present' ? Number(e.ot) || 0 : 0, remarks: str(e.remarks).trim().slice(0, 200) };
      const b = { status: normStatus(a.base[day].status || ''), ot: Number(a.base[day].ot) || 0, remarks: str(a.base[day].remarks).trim() };
      if (!same(b, c) && !same(mine, c)) bad.push(day);
    });
    if (bad.length) throw new Error('CONFLICT:' + bad.join(','));
  }
  const upd = [], add = [], del = []; let changed = 0, cleared = 0;
  const rowVals = new Array(n).fill('');
  a.entries.forEach((e) => {
    const day = Math.round(Number(e.day)); if (!(day >= 1 && day <= n)) return;
    const ds = year + '-' + pad(month + 1) + '-' + pad(day), status = checkStatus(e.status || ''), remarks = str(e.remarks).trim().slice(0, 200);
    const ot = status === 'present' ? checkOT(e.ot, status) : 0, ex = index[ds];
    if (!status && !remarks) { if (ex) { del.push(ex); cleared++; } return; }
    if (ex) upd.push({ range: q(T.att) + '!C' + ex + ':E' + ex, values: [[status, ot, remarks]] }); else add.push([worker, ds, status, ot, remarks]);
    changed++; rowVals[day - 1] = code(status, ot);
  });
  await G.batchUpdate(SHEET_ID(), upd, 'RAW');
  if (add.length) await G.append(SHEET_ID(), q(T.att) + '!A1', add, 'RAW');
  if (del.length) { const sid = (await sheetIds())[T.att]; await G.format(SHEET_ID(), del.sort((x, y) => y - x).map((r) => ({ deleteDimension: { range: { sheetId: sid, dimension: 'ROWS', startIndex: r - 1, endIndex: r } } }))); }
  await writeRegisterRange(op, { k: { year, month, days: [] } }, worker, () => ({ from: 1, vals: rowVals }));
  await logRow(op, 'MONTH EDIT', worker, year + '-' + pad(month + 1), '', changed + ' day(s) set, ' + cleared + ' cleared');
  return { ok: true, changed, cleared };
}

async function addWorker(a, op) {
  const name = checkWorker(a.name);
  const w = (await G.batchGet(SHEET_ID(), [q(T.workers) + '!A2:A2000']))[0] || [];
  if (!w.some((r) => str(r[0]).trim().toLowerCase() === name.toLowerCase())) await G.append(SHEET_ID(), q(T.workers) + '!A1', [[safe(name), 'Active']], 'RAW');
  await logRow(op, 'ADD STAFF', name, '', '', 'added');
  return { ok: true };
}
async function setStatus(a, op) {
  const name = checkWorker(a.name), active = !!a.active;
  const w = (await G.batchGet(SHEET_ID(), [q(T.workers) + '!A2:B2000']))[0] || [];
  const i = w.findIndex((r) => str(r[0]) === name); if (i < 0) throw new Error('Worker not found');
  const old = str(w[i][1]) || 'Active', now = active ? 'Active' : 'Archived';
  await G.batchUpdate(SHEET_ID(), [{ range: q(T.workers) + '!B' + (i + 2), values: [[now]] }], 'RAW');
  await logRow(op, active ? 'REACTIVATE STAFF' : 'ARCHIVE STAFF', name, '', old, now);
  return { ok: true };
}
async function rebuild(a, op) {
  const y = Math.round(Number(a.year)), m = Math.round(Number(a.month));
  if (!(y >= 2020 && y <= 2100 && m >= 0 && m <= 11)) throw new Error('Invalid month');
  return rebuildRegister(y, m, op);
}

/* ---------- Leave Plan Form ---------- */
async function leaveSubmit(a) {
  const name = str(a.name).trim().slice(0, 80), reason = str(a.reason).trim().slice(0, 500), start = str(a.start).trim(), end = str(a.end).trim();
  if (!name) throw new Error('Please enter your name.');
  if (!start || !end) throw new Error('Please pick both leave dates.');
  if (!reason) throw new Error('Please enter a reason.');
  if (!YMD.test(start)) throw new Error('Leave Start date looks invalid.');
  if (!YMD.test(end)) throw new Error('Leave End date looks invalid.');
  if (end < start) throw new Error('Leave End must be on or after Leave Start.');
  const day = (s) => Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10));
  const daysNotice = Math.round((day(start) - day(todayIST())) / 86400000), prior = daysNotice >= 90 ? 'Yes' : 'No';
  await ensureTimeFmt(T.leave);
  const ur = await G.append(SHEET_ID(), q(T.leave) + '!A1', [[nowIST(), safe(name), text(start), text(end), safe(reason), daysNotice, prior, 'Pending']], 'USER_ENTERED');
  await fmtAppended(T.leave, ur);
  const alert = await notifyLeave({ name, start, end, reason, daysNotice, appliedPrior: prior });
  return { ok: true, daysNotice, appliedPrior: prior, alert };
}
// Owner e-mail alert via a small Apps Script web app (env LEAVE_ALERT_URL / LEAVE_ALERT_SECRET). Never blocks or fails the request.
const ALERT_URL_DEFAULT = 'https://script.google.com/macros/s/AKfycbxnqHgUun_SUMuB8kOe9De86IhB5LEIgk-qBqbB-sg9Mnjv-fmfc4arC4Q7a25TUlzN/exec';
async function notifyLeave(d) {
  const secret = process.env.LEAVE_ALERT_SECRET, url = process.env.LEAVE_ALERT_URL || ALERT_URL_DEFAULT;
  if (!secret) return 'skipped';
  try {
    const ctl = new AbortController(), t = setTimeout(() => ctl.abort(), 9000);
    const r = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ ...d, secret }), signal: ctl.signal, redirect: 'follow' });
    clearTimeout(t);
    const j = await r.json().catch(() => ({}));
    return j && j.ok ? 'sent' : 'failed';
  } catch (e) { return 'failed'; }
}
async function leaveList() {
  await ensureTimeFmt(T.leave);
  const r = (await G.batchGet(SHEET_ID(), [q(T.leave) + '!B2:E5000'], 'serial'))[0] || [];
  return r.filter((x) => x[0] && x[1] && x[2]).map((x) => ({ name: str(x[0]), start: dateKey(x[1]), end: dateKey(x[2]), reason: str(x[3]) })).reverse();
}

/* ---------- Team site (public leave form on team-suryatex.vercel.app) — token-gated, see api/attendance.js ---------- */
const crypto = require('crypto');
const rowKey = (name, start, end) => crypto.createHash('sha1').update(str(name) + '|' + str(start) + '|' + str(end)).digest('hex').slice(0, 8);
async function teamLeaveList() {
  await ensureTimeFmt(T.leave);
  const r = (await G.batchGet(SHEET_ID(), [q(T.leave) + '!A2:E5000'], 'serial'))[0] || [];
  const out = [];
  r.forEach((x, i) => { if (x[1] && x[2] && x[3]) { const from = dateKey(x[2]), to = dateKey(x[3]); out.push({ id: (i + 2) + 'x' + rowKey(x[1], from, to), name: str(x[1]), from, to, reason: str(x[4]) }); } });
  return out.reverse();
}
async function teamRow(id) {   // find the sheet row for an id, refusing if that row has changed since the list was loaded
  const m = /^(\d{1,6})x([0-9a-f]{8})$/.exec(str(id)); if (!m) throw new Error('Missing entry.');
  const row = +m[1]; if (row < 2) throw new Error('Missing entry.');
  const x = ((await G.batchGet(SHEET_ID(), [q(T.leave) + '!A' + row + ':E' + row], 'serial'))[0] || [])[0] || [];
  if (!x[1] || rowKey(x[1], dateKey(x[2]), dateKey(x[3])) !== m[2]) throw new Error('This entry has changed or no longer exists. Please refresh.');
  return { row, name: str(x[1]), from: dateKey(x[2]), to: dateKey(x[3]) };
}
async function teamLeaveEdit(a, op) {
  const name = str(a.name).trim().slice(0, 80), reason = str(a.reason).trim().slice(0, 500), from = str(a.from).trim(), to = str(a.to).trim();
  if (!name) throw new Error('Please enter a name.');
  if (!YMD.test(from) || !YMD.test(to)) throw new Error('Please choose both leave dates.');
  if (to < from) throw new Error('"Leave To" cannot be before "Leave From".');
  if (!reason) throw new Error('Please enter a reason.');
  const old = await teamRow(a.id);
  await G.batchUpdate(SHEET_ID(), [{ range: q(T.leave) + '!B' + old.row + ':E' + old.row, values: [[safe(name), text(from), text(to), safe(reason)]] }], 'USER_ENTERED');
  await logRow(op, 'LEAVE EDIT', name, from, old.name + ' ' + old.from + '→' + old.to, name + ' ' + from + '→' + to);
  return { ok: true, entry: { id: old.row + 'x' + rowKey(name, from, to), name, from, to, reason } };
}
async function teamLeaveDelete(a, op) {
  const old = await teamRow(a.id), sid = (await sheetIds())[T.leave];
  if (sid == null) throw new Error('Leave sheet not found');
  await G.format(SHEET_ID(), [{ deleteDimension: { range: { sheetId: sid, dimension: 'ROWS', startIndex: old.row - 1, endIndex: old.row } } }]);
  await logRow(op, 'LEAVE DELETE', old.name, old.from, old.from + '→' + old.to, 'deleted');
  return { ok: true };
}

const ROUTES = {
  'att.getAll': () => getAll(),
  'att.save': save, 'att.bulk': bulk, 'att.monthSave': monthSave,
  'att.addWorker': addWorker, 'att.setStatus': setStatus, 'att.rebuild': rebuild,
  'leave.submit': (a) => leaveSubmit(a), 'leave.list': () => leaveList(),
  'team.leaveSubmit': (a) => leaveSubmit(a), 'team.leaveList': () => teamLeaveList(), 'team.leaveEdit': teamLeaveEdit, 'team.leaveDelete': teamLeaveDelete,
};
module.exports = { ROUTES, _t: { dateKey, normStatus, code, parse } };
