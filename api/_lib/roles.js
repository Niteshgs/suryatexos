// Roles Clarity documents (admin only). Stored on a "Roles" tab of the Users sheet, created on first use.
// A: id · B: role · C: status (Existing/Upcoming) · D: department · E: reports to · F: held by · G: purpose
// H: responsibilities · I: decides alone · J: escalates to · K: measures (KPIs) · L: notes · M: updated · N: updated by
const crypto = require('crypto');
const G = require('./google');

// Lives in the backend Users sheet by default; set ROLES_SHEET_ID to keep it in a separate Google Sheet instead.
const USERS_ID = () => process.env.ROLES_SHEET_ID || process.env.USERS_SHEET_ID || '1S8RRCmLfPbzOh8j9QxHWyV3wUy0ip1PF6SMd8V6uAdo';
const TAB = 'Roles';
const HEAD = ['ID', 'Role', 'Status', 'Department', 'Reports to', 'Held by', 'Purpose', 'Responsibilities', 'Decides alone', 'Escalates to', 'Measures', 'Notes', 'Updated', 'Updated by'];
const FIELDS = ['role', 'status', 'dept', 'reports', 'holder', 'purpose', 'duties', 'decides', 'escalate', 'kpis', 'notes'];
const STATUS = ['Existing', 'Upcoming'];
const clean = (v, n) => String(v == null ? '' : v).replace(/\r/g, '').slice(0, n || 4000);
const today = () => new Intl.DateTimeFormat('en-IN', { timeZone: 'Asia/Kolkata', day: '2-digit', month: 'short', year: 'numeric' }).format(new Date());

async function ensureTab() {
  const t = await G.sheetTitles(USERS_ID());
  if (!t.includes(TAB)) {
    await G.format(USERS_ID(), [{ addSheet: { properties: { title: TAB } } }]);
    await G.batchUpdate(USERS_ID(), [{ range: TAB + '!A1:N1', values: [HEAD] }]);
  }
}
const toObj = (r, i) => ({ row: i + 2, id: r[0], role: r[1] || '', status: r[2] || 'Existing', dept: r[3] || '', reports: r[4] || '', holder: r[5] || '', purpose: r[6] || '', duties: r[7] || '', decides: r[8] || '', escalate: r[9] || '', kpis: r[10] || '', notes: r[11] || '', updated: r[12] || '', by: r[13] || '' });
async function read() {
  try {
    const [rows] = await G.batchGet(USERS_ID(), [TAB + '!A2:N500']);
    return (rows || []).map(toObj).filter((x) => x.id && x.role);
  } catch (e) { return []; }                                   // tab not created yet
}

const ROUTES = {
  'roles.list': async () => ({ roles: await read(), statuses: STATUS }),
  'roles.save': async (a, who) => {
    const v = {}; FIELDS.forEach((k) => { v[k] = clean(a[k], k === 'role' || k === 'dept' || k === 'reports' || k === 'holder' ? 120 : 4000).trim(); });
    if (!v.role) throw new Error('Give the role a name');
    if (!STATUS.includes(v.status)) v.status = 'Existing';
    await ensureTab();
    const all = await read();
    const dup = all.find((x) => x.role.toLowerCase() === v.role.toLowerCase() && x.id !== a.id);
    if (dup) throw new Error('A role called "' + dup.role + '" already exists');
    const vals = [v.role, v.status, v.dept, v.reports, v.holder, v.purpose, v.duties, v.decides, v.escalate, v.kpis, v.notes, today(), who];
    const mine = a.id && all.find((x) => x.id === a.id);
    if (mine) {
      await G.batchUpdate(USERS_ID(), [{ range: TAB + '!B' + mine.row + ':N' + mine.row, values: [vals] }]);
      return { id: mine.id };
    }
    const id = 'r' + crypto.randomBytes(5).toString('hex');
    await G.append(USERS_ID(), TAB + '!A:N', [[id].concat(vals)]);
    return { id };
  },
  'roles.remove': async (a) => {
    const all = await read();
    const mine = all.find((x) => x.id === a.id);
    if (!mine) throw new Error('That role was not found');
    await G.clear(USERS_ID(), [TAB + '!A' + mine.row + ':N' + mine.row]);
    return { removed: true };
  },
};
module.exports = { ROUTES };
