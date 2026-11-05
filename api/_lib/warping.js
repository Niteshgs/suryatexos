// Warping: the Planner (admin / partner) and the Execute screen (staff on the warping machine), straight from Google Sheets.
// Sheet "Warping": tab "Warping Data" (the tape list), "Planning Queue" (every job, the master record) and "Execute Queue" (open jobs for the machine).
// Env: WARPING_SHEET_ID (default = the Warping workbook).
const crypto = require('crypto');
const G = require('./google');

const SHEET_ID = () => process.env.WARPING_SHEET_ID || '1Rl89_0MXlsvSdqlJi_9AsuezpGeU4BppoJj6DGlFEd8';
const DATA = 'Warping Data', PLAN = 'Planning Queue', EXEC = 'Execute Queue', LOG = 'Warping Log';   // DATA is read-only: the app never writes to it
const LOG_HEAD = ['When', 'Job ID', 'Action', 'By', 'Tape', 'Details'];
const PLAN_HEAD = ['Job ID', 'Created', 'Created by', 'Tape', 'Category', 'Stand', 'Base Ends', 'Base Yarn', 'Base Beams', 'Base Meters (each beam)',
  'Design Ends', 'Design Yarn', 'Design Beams', 'Design Meters (each beam)', 'Status', 'Started', 'Started by', 'Done', 'Done by'];
const EXEC_HEAD = ['Job ID', 'Tape', 'Stand', 'Base Yarn', 'Base Ends', 'Base Beams', 'Base Meters (each beam)',
  'Design Yarn', 'Design Ends', 'Design Beams', 'Design Meters (each beam)', 'Status', 'Started', 'Started by'];
const L = (n) => { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
const str = (v) => (v == null ? '' : String(v));
const flat = (v) => str(v).replace(/\s+/g, ' ').trim();            // tape names in the sheet contain line breaks
const safe = (s) => { s = str(s); return /^[=+\-@]/.test(s) ? "'" + s : s; };
const stamp = () => new Intl.DateTimeFormat('en-IN', { timeZone: 'Asia/Kolkata', day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit', hour12: true }).format(new Date()).replace(',', '');
const range = (tab, a, b) => G.q(tab) + '!' + a + ':' + b;
const NOACCESS = 'The app cannot open the Warping sheet. Share it with the app\'s service account as Editor.';
const wrap = async (fn) => { try { return await fn(); } catch (e) { if (/permission|403|404|not found/i.test(e.message)) throw new Error(NOACCESS); throw e; } };

async function ensureTabs() {
  const titles = await G.sheetTitles(SHEET_ID());
  const adds = [PLAN, EXEC, LOG].filter((t) => !titles.includes(t)).map((t) => ({ addSheet: { properties: { title: t } } }));
  if (adds.length) await G.format(SHEET_ID(), adds);
  const [a, b, c] = await G.batchGet(SHEET_ID(), [range(PLAN, 'A1', 'A1'), range(EXEC, 'A1', 'A1'), range(LOG, 'A1', 'A1')]);
  const data = [];
  if (!str(((a[0] || [])[0])).trim()) data.push({ range: range(PLAN, 'A1', L(PLAN_HEAD.length) + '1'), values: [PLAN_HEAD] });
  if (!str(((b[0] || [])[0])).trim()) data.push({ range: range(EXEC, 'A1', L(EXEC_HEAD.length) + '1'), values: [EXEC_HEAD] });
  if (!str(((c[0] || [])[0])).trim()) data.push({ range: range(LOG, 'A1', L(LOG_HEAD.length) + '1'), values: [LOG_HEAD] });
  if (data.length) await G.batchUpdate(SHEET_ID(), data);
}
// A log line must never block the real work
const log = (id, action, who, tape, details) => G.append(SHEET_ID(), range(LOG, 'A', L(LOG_HEAD.length)), [[stamp(), id, action, safe(who), safe(tape), safe(details || '')]]).catch(() => {});

/* ---------- tape list ---------- */
async function catalog() {
  const [rows] = await G.batchGet(SHEET_ID(), [range(DATA, 'A2', 'H')]);
  const by = new Map();
  for (const r of rows) {
    const name = flat(r[1]); if (!name) continue;
    const v = { stand: flat(r[3]) || 'Not Relevant', category: flat(r[2]), baseEnds: Number(r[4]) || 0, designEnds: Number(r[5]) || 0, baseYarn: flat(r[6]), designYarn: flat(r[7]) };
    if (!by.has(name)) by.set(name, []);
    if (!by.get(name).some((x) => x.stand === v.stand)) by.get(name).push(v);
  }
  return { tapes: [...by.entries()].map(([name, variants]) => ({ name, variants })).sort((x, y) => x.name.localeCompare(y.name, 'en', { numeric: true })) };
}

/* ---------- jobs ---------- */
const planObj = (r, i) => ({ row: i + 2, id: str(r[0]), created: str(r[1]), by: str(r[2]), tape: str(r[3]), category: str(r[4]), stand: str(r[5]),
  baseEnds: Number(r[6]) || 0, baseYarn: str(r[7]), baseBeams: Number(r[8]) || 0, baseMeters: Number(r[9]) || 0,
  designEnds: Number(r[10]) || 0, designYarn: str(r[11]), designBeams: Number(r[12]) || 0, designMeters: Number(r[13]) || 0,
  status: str(r[14]) || 'Planned', started: str(r[15]), startedBy: str(r[16]), done: str(r[17]), doneBy: str(r[18]) });
const execObj = (r, i) => ({ row: i + 2, id: str(r[0]), tape: str(r[1]), stand: str(r[2]), baseYarn: str(r[3]), baseEnds: Number(r[4]) || 0, baseBeams: Number(r[5]) || 0, baseMeters: Number(r[6]) || 0,
  designYarn: str(r[7]), designEnds: Number(r[8]) || 0, designBeams: Number(r[9]) || 0, designMeters: Number(r[10]) || 0, status: str(r[11]) || 'Planned', started: str(r[12]), startedBy: str(r[13]) });
async function readPlan() { const [rows] = await G.batchGet(SHEET_ID(), [range(PLAN, 'A2', L(PLAN_HEAD.length))]); return rows.map(planObj).filter((j) => j.id); }
async function readExec() { const [rows] = await G.batchGet(SHEET_ID(), [range(EXEC, 'A2', L(EXEC_HEAD.length))]); return rows.map(execObj).filter((j) => j.id); }

const count = (v) => { const n = Number(v); return v === '' || v == null ? 0 : (Number.isInteger(n) && n >= 0 && n <= 1000000 ? n : NaN); };
function planInputs(a) {
  const f = { baseBeams: count(a.baseBeams), baseMeters: count(a.baseMeters), designBeams: count(a.designBeams), designMeters: count(a.designMeters) };
  if (Object.values(f).some(Number.isNaN)) throw new Error('Beams and meters must be whole numbers');
  if (f.baseBeams > 99 || f.designBeams > 99) throw new Error('That is too many beams. Check the number.');
  if (!!f.baseBeams !== !!f.baseMeters) throw new Error('Base needs both the number of beams and the meters');
  if (!!f.designBeams !== !!f.designMeters) throw new Error('Design needs both the number of beams and the meters');
  if (!f.baseBeams && !f.designBeams) throw new Error('Enter the beams and meters for Base, Design, or both');
  return f;
}

async function add(a, who) {
  const name = flat(a.tape); if (!name) throw new Error('Choose a tape');
  const f = planInputs(a);
  const cat = await catalog(), t = cat.tapes.find((x) => x.name === name);
  if (!t) throw new Error('That tape is not in the Warping Data sheet');
  const v = t.variants.length === 1 ? t.variants[0] : t.variants.find((x) => x.stand === flat(a.stand));
  if (!v) throw new Error('Choose Single Stand or Double Stand');
  await ensureTabs();
  const [ids] = await G.batchGet(SHEET_ID(), [range(PLAN, 'A2', 'A')]);
  const max = ids.reduce((m, r) => { const n = Number(str(r[0]).replace(/\D/g, '')); return n > m ? n : m; }, 0);
  const id = 'W-' + String(max + 1).padStart(4, '0');
  const bEnds = f.baseBeams ? v.baseEnds : '', bYarn = f.baseBeams ? v.baseYarn : '', dEnds = f.designBeams ? v.designEnds : '', dYarn = f.designBeams ? v.designYarn : '';
  await G.append(SHEET_ID(), range(PLAN, 'A', L(PLAN_HEAD.length)), [[id, stamp(), safe(who), safe(name), safe(v.category), safe(v.stand), bEnds, safe(bYarn), f.baseBeams || '', f.baseMeters || '',
    dEnds, safe(dYarn), f.designBeams || '', f.designMeters || '', 'Planned', '', '', '', '']]);
  await G.append(SHEET_ID(), range(EXEC, 'A', L(EXEC_HEAD.length)), [[id, safe(name), safe(v.stand), safe(bYarn), bEnds, f.baseBeams || '', f.baseMeters || '', safe(dYarn), dEnds, f.designBeams || '', f.designMeters || '', 'Planned', '', '']]);
  await log(id, 'Planned', who, name, (f.baseBeams ? 'Base ' + f.baseBeams + ' x ' + f.baseMeters + ' m' : '') + (f.baseBeams && f.designBeams ? '; ' : '') + (f.designBeams ? 'Design ' + f.designBeams + ' x ' + f.designMeters + ' m' : '') + (v.stand !== 'Not Relevant' ? ' (' + v.stand + ')' : ''));
  return { id };
}

async function deleteExecRow(row) {
  const props = await G.sheetProps(SHEET_ID()), p = props.find((x) => x.title === EXEC);
  if (!p) return;
  await G.format(SHEET_ID(), [{ deleteDimension: { range: { sheetId: p.sheetId, dimension: 'ROWS', startIndex: row - 1, endIndex: row } } }]);
}
async function setStatus(a, who) {
  const want = str(a.status), id = str(a.id);
  if (!['Planned', 'Started', 'Done'].includes(want)) throw new Error('Bad status');
  const job = (await readPlan()).find((j) => j.id === id);
  if (!job) throw new Error('That job was not found');
  if (job.status === 'Cancelled') throw new Error('That job was cancelled');
  if (job.status === 'Done') throw new Error('That job is already done');
  const ok = (job.status === 'Planned' && want === 'Started') || (job.status === 'Started' && (want === 'Done' || want === 'Planned'));
  if (!ok) throw new Error(job.status === want ? 'Already ' + want.toLowerCase() : 'Start the job first');
  const at = stamp(), name = safe(who);
  const upd = want === 'Started' ? ['Started', at, name, '', ''] : want === 'Done' ? ['Done', job.started, job.startedBy, at, name] : ['Planned', '', '', '', ''];
  await G.batchUpdate(SHEET_ID(), [{ range: range(PLAN, 'O' + job.row, 'S' + job.row), values: [upd] }]);
  const ex = (await readExec()).find((j) => j.id === id);
  if (want === 'Done') { if (ex) await deleteExecRow(ex.row); }
  else if (ex) await G.batchUpdate(SHEET_ID(), [{ range: range(EXEC, 'L' + ex.row, 'N' + ex.row), values: [[want, upd[1], upd[2]]] }]);
  await log(id, want === 'Planned' ? 'Reset to Planned' : want, who, job.tape, '');
  return { status: want };
}
async function cancel(a, who) {
  const job = (await readPlan()).find((j) => j.id === str(a.id));
  if (!job) throw new Error('That job was not found');
  if (job.status === 'Done') throw new Error('A finished job cannot be cancelled');
  if (job.status === 'Cancelled') return { status: 'Cancelled' };
  await G.batchUpdate(SHEET_ID(), [{ range: range(PLAN, 'O' + job.row, 'O' + job.row), values: [['Cancelled']] }]);
  const ex = (await readExec()).find((j) => j.id === job.id);
  if (ex) await deleteExecRow(ex.row);
  await log(job.id, 'Cancelled', who, job.tape, '');
  return { status: 'Cancelled' };
}

const ROUTES = {
  'warp.catalog': () => wrap(catalog),
  'warp.list': () => wrap(async () => { try { return { jobs: (await readPlan()).filter((j) => j.status !== 'Cancelled').slice(-200) }; } catch (e) { if (/parse range|Unable to parse/i.test(e.message)) return { jobs: [] }; throw e; } }),
  'warp.add': (a, who) => wrap(() => add(a, who)),
  'warp.cancel': (a, who) => wrap(() => cancel(a, who)),
  'warp.execList': () => wrap(async () => { try { return { jobs: (await readExec()).filter((j) => j.status !== 'Done') }; } catch (e) { if (/parse range|Unable to parse/i.test(e.message)) return { jobs: [] }; throw e; } }),
  'warp.setStatus': (a, who) => wrap(() => setStatus(a, who)),
};
module.exports = { ROUTES, SHEET_ID, _t: { catalog, planInputs } };
