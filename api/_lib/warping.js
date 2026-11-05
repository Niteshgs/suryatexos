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

let TABS_OK = false;
async function ensureTabs() {
  if (TABS_OK) return;
  const titles = await G.sheetTitles(SHEET_ID());
  const adds = [PLAN, EXEC, LOG].filter((t) => !titles.includes(t)).map((t) => ({ addSheet: { properties: { title: t } } }));
  if (adds.length) await G.format(SHEET_ID(), adds);
  const [a, b, c] = await G.batchGet(SHEET_ID(), [range(PLAN, 'A1', 'A1'), range(EXEC, 'A1', 'A1'), range(LOG, 'A1', 'A1')]);
  const data = [];
  if (!str(((a[0] || [])[0])).trim()) data.push({ range: range(PLAN, 'A1', L(PLAN_HEAD.length) + '1'), values: [PLAN_HEAD] });
  if (!str(((b[0] || [])[0])).trim()) data.push({ range: range(EXEC, 'A1', L(EXEC_HEAD.length) + '1'), values: [EXEC_HEAD] });
  if (!str(((c[0] || [])[0])).trim()) data.push({ range: range(LOG, 'A1', L(LOG_HEAD.length) + '1'), values: [LOG_HEAD] });
  if (data.length) await G.batchUpdate(SHEET_ID(), data);
  TABS_OK = true;
}
// A log line must never block the real work
const log = (id, action, who, tape, details) => G.append(SHEET_ID(), range(LOG, 'A', L(LOG_HEAD.length)), [[stamp(), id, action, safe(who), safe(tape), safe(details || '')]]).catch(() => {});

/* ---------- tape list ---------- */
let CAT = { v: null, at: 0 };
async function catalog() {
  if (CAT.v && Date.now() - CAT.at < 300000) return CAT.v;
  try { CAT = { v: await catalogRead(), at: Date.now() }; return CAT.v; } catch (e) { if (quota(e) && CAT.v) return CAT.v; throw e; }
}
async function catalogRead() {
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
  designYarn: str(r[7]), designEnds: Number(r[8]) || 0, designBeams: Number(r[9]) || 0, designMeters: Number(r[10]) || 0, status: str(r[11]) || 'Ready', started: str(r[12]), startedBy: str(r[13]) });
// Google allows ~60 reads/minute for the whole app, so list screens share one short-lived copy of the Planning Queue.
// Changes (add / execute / start / done ...) always read fresh first, and clear the copy afterwards.
const PC = { rows: null, at: 0, p: null };
const quota = (e) => /429|quota/i.test(e.message || '');
function readPlan(fresh) {
  if (!fresh && PC.rows && Date.now() - PC.at < 12000) return Promise.resolve(PC.rows.slice());
  if (PC.p) return PC.p.then((r) => r.slice());
  PC.p = (async () => {
    try {
      const [rows] = await G.batchGet(SHEET_ID(), [range(PLAN, 'A2', L(PLAN_HEAD.length))]);
      PC.rows = rows.map(planObj).filter((j) => j.id); PC.at = Date.now();
      return PC.rows;
    } catch (e) { if (quota(e) && PC.rows) return PC.rows; throw e; }   // Google is busy: show the last copy instead of an error
    finally { PC.p = null; }
  })();
  return PC.p.then((r) => r.slice());
}
const dirty = () => { PC.at = 0; };
const mut = (fn) => async (a, who) => { try { return await fn(a, who); } finally { dirty(); } };   // any change clears the shared copy
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
  await log(id, 'Planned', who, name, (f.baseBeams ? 'Base ' + f.baseBeams + ' x ' + f.baseMeters + ' m' : '') + (f.baseBeams && f.designBeams ? '; ' : '') + (f.designBeams ? 'Design ' + f.designBeams + ' x ' + f.designMeters + ' m' : '') + (v.stand !== 'Not Relevant' ? ' (' + v.stand + ')' : ''));
  return { id };
}

async function deleteRow(tab, row) {
  const props = await G.sheetProps(SHEET_ID()), p = props.find((x) => x.title === tab);
  if (!p) return;
  await G.format(SHEET_ID(), [{ deleteDimension: { range: { sheetId: p.sheetId, dimension: 'ROWS', startIndex: row - 1, endIndex: row } } }]);
}
async function sendToExecute(a, who) {
  const job = (await readPlan(true)).find((j) => j.id === str(a.id));
  if (!job) throw new Error('That job was not found');
  if (job.status !== 'Planned') throw new Error(job.status === 'Cancelled' ? 'That job was cancelled' : 'Already sent to Execute');
  await ensureTabs();
  const blank = (x) => (x ? x : '');
  await G.append(SHEET_ID(), range(EXEC, 'A', L(EXEC_HEAD.length)), [[job.id, safe(job.tape), safe(job.stand), safe(job.baseYarn), blank(job.baseEnds), blank(job.baseBeams), blank(job.baseMeters),
    safe(job.designYarn), blank(job.designEnds), blank(job.designBeams), blank(job.designMeters), 'Ready', '', '']]);
  await G.batchUpdate(SHEET_ID(), [{ range: range(PLAN, 'O' + job.row, 'O' + job.row), values: [['Ready']] }]);
  await log(job.id, 'Sent to Execute', who, job.tape, '');
  return { status: 'Ready' };
}
async function setStatus(a, who) {
  const want = str(a.status), id = str(a.id);
  if (!['Started', 'Done', 'Ready', 'Hold'].includes(want)) throw new Error('Bad status');
  const job = (await readPlan(true)).find((j) => j.id === id);
  if (!job) throw new Error('That job was not found');
  if (job.status === 'Cancelled') throw new Error('That job was cancelled');
  if (job.status === 'Done' || job.status === 'Archived') throw new Error('That job is already done');
  if (job.status === 'Planned') throw new Error('That job has not been sent to Execute yet');
  const next = { Ready: ['Started'], Started: ['Done', 'Hold', 'Ready'], 'On hold': ['Started'] }[job.status] || [];
  if (!next.includes(want)) throw new Error(job.status === want || (job.status === 'On hold' && want === 'Hold') ? 'Already ' + job.status.toLowerCase() : job.status === 'Ready' ? 'Start the job first' : 'That is not possible right now');
  const at = stamp(), name = safe(who), resume = job.status === 'On hold';
  const upd = want === 'Started' ? ['Started', resume && job.started ? job.started : at, resume && job.startedBy ? job.startedBy : name, '', '']
    : want === 'Done' ? ['Done', job.started, job.startedBy, at, name]
    : want === 'Hold' ? ['On hold', job.started, job.startedBy, '', ''] : ['Ready', '', '', '', ''];
  await G.batchUpdate(SHEET_ID(), [{ range: range(PLAN, 'O' + job.row, 'S' + job.row), values: [upd] }]);
  const ex = (await readExec()).find((j) => j.id === id);
  if (want === 'Done') { if (ex) await deleteRow(EXEC, ex.row); }
  else if (ex) await G.batchUpdate(SHEET_ID(), [{ range: range(EXEC, 'L' + ex.row, 'N' + ex.row), values: [[upd[0], upd[1], upd[2]]] }]);
  await log(id, want === 'Ready' ? 'Reset to Ready' : want === 'Hold' ? 'Put on hold' : resume ? 'Resumed' : want, who, job.tape, '');
  return { status: upd[0] };
}
async function archive(a, who, back) {
  const job = (await readPlan(true)).find((j) => j.id === str(a.id));
  if (!job) throw new Error('That job was not found');
  const from = back ? 'Archived' : 'Done', to = back ? 'Done' : 'Archived';
  if (job.status === to) return { status: to };
  if (job.status !== from) throw new Error(back ? 'Only an archived job can be restored' : 'Only a finished job can be archived');
  await G.batchUpdate(SHEET_ID(), [{ range: range(PLAN, 'O' + job.row, 'O' + job.row), values: [[to]] }]);
  await log(job.id, back ? 'Restored from archive' : 'Archived', who, job.tape, '');
  return { status: to };
}
/* ---------- tape photo (Warping Data column A, read only) ---------- */
const IMG = new Map();
async function image(a) {
  const name = flat(a.tape); if (!name) throw new Error('No tape given');
  const hit = IMG.get(name.toLowerCase());
  if (hit && Date.now() - hit.at < 600000) return { src: hit.src };
  // In-cell pictures are not exposed by the Sheets REST API, so a tiny read-only Apps Script (see the Apps Script given to the owner) hands them over.
  const hook = process.env.WARPING_IMG_URL, key = process.env.WARPING_IMG_KEY;
  if (!hook || !key) throw new Error('Tape photos are not connected yet');
  const r = await fetch(hook + (hook.includes('?') ? '&' : '?') + 'key=' + encodeURIComponent(key) + '&tape=' + encodeURIComponent(name), { redirect: 'follow' });
  const tx = await r.text(); let j; try { j = JSON.parse(tx); } catch (e) { throw new Error('Photo service did not answer properly (' + r.status + ': ' + tx.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 140) + ')'); }
  if (!j || j.error) throw new Error((j && j.error) || 'Could not load the photo');
  if (!j.src) throw new Error('No photo for this tape in the sheet');
  if (IMG.size > 30) IMG.clear();
  IMG.set(key0(name), { src: j.src, at: Date.now() });
  return { src: j.src };
}
const key0 = (n) => n.toLowerCase();
/* ---------- analytics: time from Start to Done (hold time taken out) ---------- */
const MON = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
function parseStamp(t, nowMs) {                       // "5 Nov 9:33 pm" (India time, no year) -> ms; the year is the latest one that is not in the future
  const m = /^(\d{1,2})\s+([A-Za-z]{3})[a-z]*\.?\s+(\d{1,2}):(\d{2})\s*([ap])\.?m/i.exec(str(t).trim());
  if (!m || MON[m[2].toLowerCase()] === undefined) return NaN;
  let h = +m[3] % 12 + (m[5].toLowerCase() === 'p' ? 12 : 0);
  const yNow = new Date(nowMs + 19800000).getUTCFullYear();
  for (let y = yNow; y >= yNow - 1; y--) {
    const ms = Date.UTC(y, MON[m[2].toLowerCase()], +m[1], h, +m[4]) - 19800000;
    if (ms <= nowMs + 86400000) return ms;
  }
  return NaN;
}
async function analytics() {
  const now = Date.now();
  let plan = []; try { plan = await readPlan(); } catch (e) { if (!/parse range|Unable to parse/i.test(e.message)) throw e; }
  let logRows = []; try { [logRows] = await G.batchGet(SHEET_ID(), [range(LOG, 'A2', 'C')]); } catch (e) { if (!/parse range|Unable to parse/i.test(e.message)) throw e; }
  const ev = new Map();
  for (const r of logRows) {
    const act = str(r[2]); if (act !== 'Put on hold' && act !== 'Resumed') continue;
    const t = parseStamp(r[0], now); if (Number.isNaN(t)) continue;
    const id = str(r[1]); if (!ev.has(id)) ev.set(id, []);
    ev.get(id).push({ t, act });
  }
  const jobs = [];
  for (const j of plan) {
    if (j.status !== 'Done' && j.status !== 'Archived') continue;
    const s = parseStamp(j.started, now), d = parseStamp(j.done, now);
    if (Number.isNaN(s) || Number.isNaN(d) || d < s) continue;
    let hold = 0, since = null;
    for (const e of (ev.get(j.id) || []).sort((x, y) => x.t - y.t)) {
      if (e.t < s || e.t > d) continue;
      if (e.act === 'Put on hold') { if (since === null) since = e.t; }
      else if (since !== null) { hold += e.t - since; since = null; }
    }
    if (since !== null) hold += d - since;
    const beams = j.baseBeams + j.designBeams, meters = j.baseBeams * j.baseMeters + j.designBeams * j.designMeters;
    const elapsed = (d - s) / 60000, held = Math.min(hold / 60000, elapsed), net = elapsed - held;
    if (!beams) continue;
    jobs.push({ id: j.id, tape: j.tape, category: j.category, stand: j.stand, beams, meters, started: j.started, done: j.done, doneMs: d,
      elapsedMin: Math.round(elapsed), holdMin: Math.round(held), netMin: Math.round(net), perBeamMin: Math.round(net / beams), metersPerHour: net > 0 ? Math.round(meters / (net / 60)) : 0 });
  }
  const agg = (list) => {
    const beams = list.reduce((a, j) => a + j.beams, 0), net = list.reduce((a, j) => a + j.netMin, 0), meters = list.reduce((a, j) => a + j.meters, 0);
    return { jobs: list.length, beams, meters, avgPerBeamMin: beams ? Math.round(net / beams) : 0, avgJobMin: list.length ? Math.round(net / list.length) : 0, metersPerHour: net > 0 ? Math.round(meters / (net / 60)) : 0 };
  };
  const groups = (key) => { const m = new Map(); for (const j of jobs) { const k = key(j); if (!k) continue; if (!m.has(k)) m.set(k, []); m.get(k).push(j); } return [...m.entries()].map(([name, l]) => ({ name, ...agg(l) })); };
  return {
    overall: agg(jobs),
    byCategory: groups((j) => j.category).sort((a, b) => a.name.localeCompare(b.name)),
    byTape: groups((j) => j.tape).sort((a, b) => b.jobs - a.jobs || a.name.localeCompare(b.name)).slice(0, 40),
    recent: jobs.sort((a, b) => b.doneMs - a.doneMs).slice(0, 30).map(({ doneMs, ...r }) => r),
  };
}
async function addTape(a, who) {
  const name = flat(a.name);
  if (!name) throw new Error('Enter the tape name');
  if (name.length > 80) throw new Error('That tape name is too long');
  const category = a.category === 'Knitted' ? 'Knitted' : a.category === 'Woven' ? 'Woven' : '';
  if (!category) throw new Error('Choose Woven or Knitted');
  const stand = category === 'Knitted' ? 'Not Relevant' : ['Single Stand', 'Double Stand'].includes(a.stand) ? a.stand : '';
  if (!stand) throw new Error('Choose Single Stand or Double Stand');
  const be = count(a.baseEnds), de = count(a.designEnds);
  if (Number.isNaN(be) || Number.isNaN(de) || be > 100000 || de > 100000) throw new Error('Ends must be whole numbers');
  if (!be) throw new Error('Enter the Base ends');
  const by = flat(a.baseYarn).slice(0, 60), dy = flat(a.designYarn).slice(0, 60);
  const [rows] = await G.batchGet(SHEET_ID(), [range(DATA, 'A2', 'H')]);
  if (rows.some((r) => flat(r[1]).toLowerCase() === name.toLowerCase() && (flat(r[3]) || 'Not Relevant') === stand)) throw new Error('That tape and stand is already in the sheet');
  const row = rows.length + 2;
  await G.batchUpdate(SHEET_ID(), [{ range: range(DATA, 'B' + row, 'H' + row), values: [[safe(name), category, stand, be, de, safe(by), safe(dy)]] }]);
  CAT = { v: null, at: 0 };
  await ensureTabs();
  await log('', 'Tape added', who, name, category + ' · ' + stand + ' · ' + be + '/' + de + ' ends');
  return { added: true };
}
async function edit(a, who) {
  const job = (await readPlan(true)).find((j) => j.id === str(a.id));
  if (!job) throw new Error('That job was not found');
  if (job.status !== 'Planned') throw new Error('Only a plan that has not gone to Execute can be edited');
  const f = planInputs(a);
  const cat = await catalog(), t = cat.tapes.find((x) => x.name === flat(job.tape));
  const v = t && (t.variants.find((x) => x.stand === job.stand) || (t.variants.length === 1 ? t.variants[0] : null));
  if (!v) throw new Error('That tape is no longer in the Warping Data sheet');
  await G.batchUpdate(SHEET_ID(), [{ range: range(PLAN, 'G' + job.row, 'N' + job.row), values: [[
    f.baseBeams ? v.baseEnds : '', f.baseBeams ? safe(v.baseYarn) : '', f.baseBeams || '', f.baseMeters || '',
    f.designBeams ? v.designEnds : '', f.designBeams ? safe(v.designYarn) : '', f.designBeams || '', f.designMeters || '']] }]);
  await log(job.id, 'Plan edited', who, job.tape, (f.baseBeams ? 'Base ' + f.baseBeams + ' x ' + f.baseMeters + ' m' : '') + (f.baseBeams && f.designBeams ? '; ' : '') + (f.designBeams ? 'Design ' + f.designBeams + ' x ' + f.designMeters + ' m' : ''));
  return { edited: true };
}
async function remove(a, who) {
  const job = (await readPlan(true)).find((j) => j.id === str(a.id));
  if (!job) throw new Error('That job was not found');
  if (job.status !== 'Planned') throw new Error('Only a plan that has not gone to Execute can be deleted. Cancel it instead.');
  await deleteRow(PLAN, job.row);
  await log(job.id, 'Plan deleted', who, job.tape, '');
  return { removed: true };
}
async function cancel(a, who) {
  const job = (await readPlan(true)).find((j) => j.id === str(a.id));
  if (!job) throw new Error('That job was not found');
  if (job.status === 'Done' || job.status === 'Archived') throw new Error('A finished job cannot be cancelled');
  if (job.status === 'Cancelled') return { status: 'Cancelled' };
  await G.batchUpdate(SHEET_ID(), [{ range: range(PLAN, 'O' + job.row, 'O' + job.row), values: [['Cancelled']] }]);
  const ex = (await readExec()).find((j) => j.id === job.id);
  if (ex) await deleteRow(EXEC, ex.row);
  await log(job.id, 'Cancelled', who, job.tape, '');
  return { status: 'Cancelled' };
}

const ROUTES = {
  'warp.catalog': () => wrap(catalog),
  'warp.list': () => wrap(async () => { try { return { jobs: (await readPlan()).filter((j) => j.status !== 'Cancelled' && j.status !== 'Done').slice(-200) }; } catch (e) { if (/parse range|Unable to parse/i.test(e.message)) return { jobs: [] }; throw e; } }),
  'warp.add': (a, who) => wrap(() => mut(add)(a, who)),
  'warp.toExecute': (a, who) => wrap(() => mut(sendToExecute)(a, who)),
  'warp.summary': () => wrap(async () => {
    let jobs = [];
    try { jobs = await readPlan(); } catch (e) { if (!/parse range|Unable to parse/i.test(e.message)) throw e; }
    const today = new Intl.DateTimeFormat('en-IN', { timeZone: 'Asia/Kolkata', day: 'numeric', month: 'short' }).format(new Date()).replace(',', '') + ' ';
    const n = (f) => jobs.filter(f).length;
    return { planning: n((j) => j.status === 'Planned'), execute: n((j) => ['Ready', 'Started', 'On hold'].includes(j.status)), hold: n((j) => j.status === 'On hold'), doneToday: n((j) => (j.status === 'Done' || j.status === 'Archived') && str(j.done).startsWith(today)) };
  }),
  'warp.board': () => wrap(async () => { try { return { jobs: (await readPlan()).filter((j) => j.status !== 'Cancelled').reverse().slice(0, 150) }; } catch (e) { if (/parse range|Unable to parse/i.test(e.message)) return { jobs: [] }; throw e; } }),
  'warp.archive': (a, who) => wrap(() => mut((x, w) => archive(x, w, false))(a, who)),
  'warp.unarchive': (a, who) => wrap(() => mut((x, w) => archive(x, w, true))(a, who)),
  'warp.archived': () => wrap(async () => { try { return { jobs: (await readPlan()).filter((j) => j.status === 'Archived').reverse().slice(0, 150) }; } catch (e) { if (/parse range|Unable to parse/i.test(e.message)) return { jobs: [] }; throw e; } }),
  'warp.ready': () => wrap(async () => { try { return { jobs: (await readPlan()).filter((j) => j.status === 'Done').reverse().slice(0, 100) }; } catch (e) { if (/parse range|Unable to parse/i.test(e.message)) return { jobs: [] }; throw e; } }),
  'warp.image': (a) => image(a),
  'warp.analytics': () => wrap(analytics),
  'warp.addTape': (a, who) => wrap(() => addTape(a, who)),
  'warp.edit': (a, who) => wrap(() => mut(edit)(a, who)),
  'warp.remove': (a, who) => wrap(() => mut(remove)(a, who)),
  'warp.cancel': (a, who) => wrap(() => mut(cancel)(a, who)),
  'warp.execList': () => wrap(async () => { try { return { jobs: (await readPlan()).filter((j) => ['Ready', 'Started', 'On hold'].includes(j.status)) }; } catch (e) { if (/parse range|Unable to parse/i.test(e.message)) return { jobs: [] }; throw e; } }),
  'warp.setStatus': (a, who) => wrap(() => mut(setStatus)(a, who)),
};
module.exports = { ROUTES, SHEET_ID, _t: { catalog, planInputs } };
