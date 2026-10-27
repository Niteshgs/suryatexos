// Sign-in and permissions for Suryatex OS.
//  - Google Sign-In: the browser sends a Google ID token; we verify its signature against Google's public keys.
//  - People and what they may open live in the "Suryatex OS Users" Google Sheet (read/written via the service account).
//  - A signed, HttpOnly session cookie remembers who is signed in. Every /api route re-checks the sheet,
//    so un-ticking a box or setting Active = No takes effect within about half a minute.
// Env (all optional, sensible defaults): GOOGLE_CLIENT_ID, USERS_SHEET_ID, SESSION_SECRET
const crypto = require('crypto');
const G = require('./google');

const CLIENT_ID = () => process.env.GOOGLE_CLIENT_ID || '788054844958-g97rekd42omi2rpe0r671l0ehh766un1.apps.googleusercontent.com';
const USERS_ID = () => process.env.USERS_SHEET_ID || '1S8RRCmLfPbzOh8j9QxHWyV3wUy0ip1PF6SMd8V6uAdo';
const COOKIE = 'sx_session';
const MAX_AGE = 30 * 24 * 3600;           // stay signed in for 30 days
const ROLES = ['Admin', 'Partner', 'Staff'];

// Everything a person can be given. Add new functions here as they are built; the admin screen picks them up.
const CATALOG = [
  { id: 'yarn', name: 'Yarn Stock', items: [
    { id: 'yarn.dash', name: 'Dashboard' }, { id: 'yarn.entry', name: 'Material entry' }, { id: 'yarn.proc', name: 'Orders' } ] },
  { id: 'fg', name: 'Finished Goods', items: [
    { id: 'fg.status', name: 'Stock status' }, { id: 'fg.po', name: 'PO Track' }, { id: 'fg.sweep', name: 'Stock Sweep' } ] },
  { id: 'knitted', name: 'Knitted Production', items: [
    { id: 'knitted.entry', name: 'Entry' }, { id: 'knitted.dash', name: 'Dashboard' }, { id: 'knitted.admin', name: 'Admin (tapes & machine activity)' } ] },
  { id: 'sales', name: 'Sales', items: [
    { id: 'sales.touch', name: 'Customer Touch Points (view & tick)' }, { id: 'sales.touchmanage', name: 'Touch Points: add, edit & archive customers' } ] },
  { id: 'expenses', name: 'Cash Expenses', items: [
    { id: 'expenses.add', name: 'Add expenses' }, { id: 'expenses.view', name: 'View expense list & monthly totals' } ] },
  { id: 'attendance', name: 'Attendance', items: [
    { id: 'attendance.mark', name: 'Mark daily attendance' }, { id: 'attendance.view', name: 'Register & Profiles (view)' },
    { id: 'attendance.edit', name: 'Edit full month & Bulk Mark' }, { id: 'attendance.staff', name: 'Staff list: add, archive, rebuild Register' } ] },
  { id: 'leave', name: 'Leave Request', items: [
    { id: 'attendance.leave', name: 'Leave Request (apply & see who applied)' } ] },
  { id: 'calc', name: 'Calculators', items: [
    { id: 'calc.payment', name: 'Payment Calculator' }, { id: 'calc.gross', name: 'Gross Value Calculator' } ] },
];
const ALL_KEYS = CATALOG.flatMap((g) => g.items.map((i) => i.id));

// Which permission opens which API call (any one of the listed keys is enough).
const ROUTE_PERMS = {
  yarn: {
    getAll: ['yarn.dash', 'yarn.entry', 'yarn.proc'],
    saveStock: ['yarn.entry'], updateMin: ['yarn.entry'], addMaterial: ['yarn.entry'],
    addOrder: ['yarn.proc'], updateOrder: ['yarn.dash', 'yarn.proc'],
  },
  fg: { getAll: ['fg.status', 'fg.po', 'fg.sweep'], sweepPhotos: ['fg.sweep'] },
  hub: {
    'knitted.getAll': ['knitted.entry', 'knitted.dash', 'knitted.admin'],
    'knitted.saveBeams': ['knitted.entry'],
    'knitted.verifyPin': ['knitted.admin'], 'knitted.addTape': ['knitted.admin'], 'knitted.removeTape': ['knitted.admin'],
    'knitted.renameTape': ['knitted.admin'], 'knitted.activity': ['knitted.admin'],
  },
  sales: {
    'touch.getAll': ['sales.touch'], 'touch.toggle': ['sales.touch'],
    'touch.addClient': ['sales.touchmanage'], 'touch.updateClient': ['sales.touchmanage'], 'touch.archiveClient': ['sales.touchmanage'],
  },
  attendance: {
    'att.getAll': ['attendance.mark', 'attendance.view', 'attendance.edit', 'attendance.staff'], 'att.save': ['attendance.mark'],
    'att.bulk': ['attendance.edit'], 'att.monthSave': ['attendance.edit'],
    'att.addWorker': ['attendance.staff'], 'att.setStatus': ['attendance.staff'], 'att.rebuild': ['attendance.staff'],
    'leave.submit': ['attendance.leave'], 'leave.list': ['attendance.leave'],
  },
  expenses: { 'expenses.add': ['expenses.add'], 'expenses.list': ['expenses.view'], 'expenses.totals': ['expenses.view'] },
};

/* ---------- small helpers ---------- */
const b64u = (b) => Buffer.from(b).toString('base64url');
function secret() {
  if (process.env.SESSION_SECRET) return process.env.SESSION_SECRET;
  const raw = process.env.GOOGLE_SERVICE_KEY;                                  // derived, so no extra secret has to be stored
  return raw ? crypto.createHash('sha256').update('sx-session|' + raw).digest('hex') : null;
}
function sign(payload) {
  const s = secret(); if (!s) throw new Error('Sign-in is not configured');
  const body = b64u(JSON.stringify(payload));
  return body + '.' + crypto.createHmac('sha256', s).update(body).digest('base64url');
}
function unsign(tok) {
  const s = secret(); if (!s || typeof tok !== 'string') return null;
  const [body, mac] = tok.split('.'); if (!body || !mac) return null;
  const good = crypto.createHmac('sha256', s).update(body).digest('base64url');
  const a = Buffer.from(mac), b = Buffer.from(good);
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return null;
  try { return JSON.parse(Buffer.from(body, 'base64url').toString()); } catch (e) { return null; }
}
function cookies(req) {
  const out = {};
  String((req.headers && req.headers.cookie) || '').split(';').forEach((p) => {
    const i = p.indexOf('='); if (i > 0) out[p.slice(0, i).trim()] = p.slice(i + 1).trim();
  });
  return out;
}
function setSession(res, email) {
  const tok = sign({ e: email, x: Math.floor(Date.now() / 1000) + MAX_AGE });
  res.setHeader('Set-Cookie', COOKIE + '=' + tok + '; Path=/; Max-Age=' + MAX_AGE + '; HttpOnly; Secure; SameSite=Lax');
}
function clearSession(res) {
  res.setHeader('Set-Cookie', COOKIE + '=; Path=/; Max-Age=0; HttpOnly; Secure; SameSite=Lax');
}

/* ---------- Google ID token ---------- */
let certs = { keys: null, exp: 0 };
async function googleKeys(force) {
  if (!force && certs.keys && Date.now() < certs.exp) return certs.keys;
  const r = await fetch('https://www.googleapis.com/oauth2/v3/certs');
  const j = await r.json();
  certs = { keys: j.keys || [], exp: Date.now() + 3600 * 1000 };
  return certs.keys;
}
async function verifyGoogle(idToken) {
  const bad = (m) => Object.assign(new Error(m || 'Google sign-in could not be verified'), { code: 'BADTOKEN' });
  const parts = String(idToken || '').split('.');
  if (parts.length !== 3) throw bad();
  let header, claims;
  try { header = JSON.parse(Buffer.from(parts[0], 'base64url').toString()); claims = JSON.parse(Buffer.from(parts[1], 'base64url').toString()); } catch (e) { throw bad(); }
  if (header.alg !== 'RS256') throw bad();
  let jwk = (await googleKeys()).find((k) => k.kid === header.kid);
  if (!jwk) jwk = (await googleKeys(true)).find((k) => k.kid === header.kid);
  if (!jwk) throw bad();
  const ok = crypto.verify('RSA-SHA256', Buffer.from(parts[0] + '.' + parts[1]), crypto.createPublicKey({ key: jwk, format: 'jwk' }), Buffer.from(parts[2], 'base64url'));
  if (!ok) throw bad();
  const now = Math.floor(Date.now() / 1000);
  if (claims.aud !== CLIENT_ID()) throw bad();
  if (claims.iss !== 'https://accounts.google.com' && claims.iss !== 'accounts.google.com') throw bad();
  if (!(claims.exp > now) || claims.iat > now + 300) throw bad('Google sign-in expired, please try again');
  if (!claims.email || claims.email_verified !== true) throw bad();
  return { email: String(claims.email).trim().toLowerCase(), name: claims.name || '' };
}

/* ---------- Users sheet ---------- */
const HEADER = ['Email', 'Name', 'Role', 'Active', 'Access', 'Last sign-in'];
let cache = { at: 0, list: null };
const yes = (v) => /^(yes|y|true|1|active)$/i.test(String(v == null ? '' : v).trim());
function roleOf(v) { const r = String(v || '').trim().toLowerCase(); return r === 'admin' ? 'Admin' : r === 'partner' ? 'Partner' : 'Staff'; }
function keysOf(v) {
  const seen = new Set();
  String(v || '').split(/[\s,;]+/).forEach((k) => { if (ALL_KEYS.includes(k)) seen.add(k); });
  return [...seen];
}
function parseUsers(rows) {
  const head = (rows[0] || []).map((h) => String(h).trim().toLowerCase());
  const col = (name, dflt) => { const i = head.indexOf(name); return i >= 0 ? i : dflt; };
  const c = { email: col('email', 0), name: col('name', 1), role: col('role', 2), active: col('active', 3), access: col('access', 4), last: col('last sign-in', 5) };
  const out = [];
  rows.slice(1).forEach((r, i) => {
    const email = String(r[c.email] || '').trim().toLowerCase();
    if (!email) return;
    out.push({ row: i + 2, email, name: String(r[c.name] || '').trim(), role: roleOf(r[c.role]), active: yes(r[c.active]),
      access: keysOf(r[c.access]), lastSignIn: String(r[c.last] || '') });
  });
  return out;
}
async function loadUsers(force) {
  if (!force && cache.list && Date.now() - cache.at < 30000) return cache.list;
  try {
    const [rows] = await G.batchGet(USERS_ID(), ['A1:F500']);
    cache = { at: Date.now(), list: parseUsers(rows) };
  } catch (e) {
    if (cache.list && Date.now() - cache.at < 3600 * 1000) return cache.list;   // brief Google hiccup: keep using the last list
    throw e;
  }
  return cache.list;
}
const forget = () => { cache = { at: 0, list: null }; };

function meOf(u) {
  const isAdmin = u.role === 'Admin';
  return { email: u.email, name: u.name || u.email.split('@')[0], role: u.role, isAdmin, access: isAdmin ? ALL_KEYS.slice() : u.access };
}

/* ---------- request checks ---------- */
async function whoami(req) {
  const p = unsign(cookies(req)[COOKIE]);
  if (!p || !(p.x > Date.now() / 1000)) return null;
  const u = (await loadUsers()).find((x) => x.email === p.e);
  return u && u.active ? meOf(u) : null;
}
function sameOrigin(req) {
  const o = req.headers && req.headers.origin; if (!o) return true;
  try { return new URL(o).host === req.headers.host; } catch (e) { return false; }
}
// Signed-in check. Sends the error response itself and returns null when the caller must stop.
async function guard(req, res) {
  if (!sameOrigin(req)) { res.status(403).json({ ok: false, error: 'Blocked' }); return null; }
  let me;
  try { me = await whoami(req); } catch (e) { res.status(503).json({ ok: false, error: 'Could not check your access right now. Please try again.' }); return null; }
  if (!me) { res.status(401).json({ ok: false, auth: true, error: 'Please sign in' }); return null; }
  return me;
}
// Permission check for one call: any one of `keys` is enough ('admin' = admins only).
function allow(me, res, keys) {
  const ok = keys === 'admin' ? me.isAdmin : Array.isArray(keys) && keys.some((k) => me.access.includes(k));
  if (!ok) res.status(403).json({ ok: false, forbidden: true, error: 'You do not have access to this' });
  return ok;
}

/* ---------- admin: people ---------- */
const stamp = () => new Intl.DateTimeFormat('en-IN', { timeZone: 'Asia/Kolkata', day: '2-digit', month: 'short', year: 'numeric', hour: 'numeric', minute: '2-digit', hour12: true }).format(new Date()).replace(',', '');
async function touch(u) {
  try { await G.batchUpdate(USERS_ID(), [{ range: 'F' + u.row, values: [[stamp()]] }]); u.lastSignIn = stamp(); } catch (e) { /* not important */ }
}
async function ensureHeader() {
  const [rows] = await G.batchGet(USERS_ID(), ['A1:F1']);
  const have = (rows[0] || []).map((h) => String(h).trim().toLowerCase());
  if (HEADER.some((h) => !have.includes(h.toLowerCase()))) await G.batchUpdate(USERS_ID(), [{ range: 'A1:F1', values: [HEADER] }]);
}
async function listUsers() {
  const list = await loadUsers(true);
  return list.map((u) => ({ email: u.email, name: u.name, role: u.role, active: u.active, access: u.role === 'Admin' ? ALL_KEYS.slice() : u.access, lastSignIn: u.lastSignIn }));
}
async function saveUser(me, input) {
  const email = String((input && input.email) || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Enter a valid email address');
  const role = ROLES.includes(input.role) ? input.role : 'Staff';
  const active = input.active !== false;
  const name = String(input.name || '').trim().slice(0, 60);
  const access = role === 'Admin' ? [] : keysOf((Array.isArray(input.access) ? input.access : []).join(','));
  const list = await loadUsers(true);
  const existing = list.find((u) => u.email === email);
  if (email === me.email && (role !== 'Admin' || !active)) throw new Error("You can't remove your own admin access");
  const admins = list.filter((u) => u.active && u.role === 'Admin' && u.email !== email).length + (role === 'Admin' && active ? 1 : 0);
  if (admins < 1) throw new Error('There must be at least one active admin');
  await ensureHeader();
  const row = [email, name, role, active ? 'Yes' : 'No', access.join(', ')];
  if (existing) await G.batchUpdate(USERS_ID(), [{ range: 'A' + existing.row + ':E' + existing.row, values: [row] }]);
  else await G.append(USERS_ID(), 'A:F', [row]);
  forget();
  return listUsers();
}

module.exports = { CATALOG, ALL_KEYS, ROUTE_PERMS, ROLES, verifyGoogle, loadUsers, meOf, setSession, clearSession, whoami, guard, allow, touch, listUsers, saveUser, forget };
