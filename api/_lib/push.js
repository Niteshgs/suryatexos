// Phone notifications (Web Push) for Suryatex OS. Subscriptions live on a "Push" tab of the Users sheet:
//   A: email · B: subscription JSON · C: added · D: role (Admin/Partner) · (E1: the date the "attendance complete" alert last went out)
// The push key pair is derived from the server secret, so there is nothing extra to store or leak.
const crypto = require('crypto');
const G = require('./google');
const webpush = require('web-push');

const USERS_ID = () => process.env.USERS_SHEET_ID || '1S8RRCmLfPbzOh8j9QxHWyV3wUy0ip1PF6SMd8V6uAdo';
const TAB = 'Push';
const b64u = (b) => Buffer.from(b).toString('base64url');

let keys = null;
function vapid() {
  if (keys) return keys;
  const s = process.env.SESSION_SECRET || process.env.GOOGLE_SERVICE_KEY;
  if (!s) throw new Error('Notifications are not configured');
  const priv = crypto.createHash('sha256').update('sx-vapid|' + s).digest();
  const ecdh = crypto.createECDH('prime256v1'); ecdh.setPrivateKey(priv);
  keys = { publicKey: b64u(ecdh.getPublicKey()), privateKey: b64u(priv) };
  webpush.setVapidDetails('mailto:noreply@suryatexos.vercel.app', keys.publicKey, keys.privateKey);
  return keys;
}
const publicKey = () => vapid().publicKey;

async function ensureTab() {
  const t = await G.sheetTitles(USERS_ID());
  if (!t.includes(TAB)) {
    await G.format(USERS_ID(), [{ addSheet: { properties: { title: TAB } } }]);
    await G.batchUpdate(USERS_ID(), [{ range: TAB + '!A1:D1', values: [['Email', 'Subscription (do not edit)', 'Added', 'Role']] }]);
  }
}
async function read() {
  try {
    const [rows, meta] = await G.batchGet(USERS_ID(), [TAB + '!A2:D500', TAB + '!E1:E1']);
    const subs = [];
    rows.forEach((r, i) => { if (r[0] && r[1]) { try { subs.push({ row: i + 2, email: String(r[0]).toLowerCase(), role: String(r[3] || 'Admin'), sub: JSON.parse(r[1]) }); } catch (e) {} } });
    return { subs, last: String(((meta || [])[0] || [])[0] || '') };
  } catch (e) { return { subs: [], last: '', missing: true }; }
}

async function subscribe(email, role, sub) {
  if (!sub || !sub.endpoint || !sub.keys) throw new Error('Bad subscription');
  vapid(); await ensureTab();
  const { subs } = await read();
  const mine = subs.find((s) => s.sub.endpoint === sub.endpoint);
  const row = [email.toLowerCase(), JSON.stringify(sub), new Date().toISOString().slice(0, 10), role];
  if (mine) await G.batchUpdate(USERS_ID(), [{ range: TAB + '!A' + mine.row + ':D' + mine.row, values: [row] }]);
  else await G.append(USERS_ID(), TAB + '!A:D', [row]);
  cache = { at: 0 };
  return { on: true };
}
async function unsubscribe(endpoint) {
  const { subs } = await read();
  const mine = subs.filter((s) => s.sub.endpoint === endpoint);
  if (mine.length) await G.clear(USERS_ID(), mine.map((s) => TAB + '!A' + s.row + ':D' + s.row));
  cache = { at: 0 };
  return { on: false };
}
async function status(endpoint) {
  const { subs } = await read();
  return { on: !!endpoint && subs.some((s) => s.sub.endpoint === endpoint), key: publicKey() };
}

let cache = { at: 0 };                       // so a burst of saves reads the Push tab at most once a minute
async function meta() { if (cache.data && Date.now() - cache.at < 60000) return cache.data; const d = await read(); cache = { at: Date.now(), data: d }; return d; }

// Is there anyone who could still be told today? (cheap, cached) — lets saves skip all the extra work after the alert has gone out.
async function pending(today) { const m = await meta(); return m.subs.length > 0 && m.last !== today; }

// Who hears about it: a Partner finishing -> only Admins; Staff (or in-charge) finishing -> Admins and Partners; an Admin finishing -> Partners.
function audience(subs, actor) {
  const r = actor && actor.role;
  return subs.filter((s) => s.email !== String(actor && actor.email || '').toLowerCase() && (r === 'Partner' ? s.role === 'Admin' : r === 'Admin' ? s.role === 'Partner' : true));
}

// Called when a save has just left nobody unmarked today. Sends ONE alert per day.
async function attendanceDone(today, counts, actor) {
  cache = { at: 0 };
  const fresh = await read();                                              // re-check right before claiming (another save may have won)
  if (fresh.last === today || !fresh.subs.length) return false;
  await G.batchUpdate(USERS_ID(), [{ range: TAB + '!E1', values: [[today]] }], 'RAW');   // claim first, so it can never go out twice
  cache = { at: Date.now(), data: Object.assign({}, fresh, { last: today }) };
  vapid();
  const body = counts.present + ' present · ' + counts.absent + ' absent' + (counts.total ? ' · ' + counts.total + ' staff' : '');
  const payload = JSON.stringify({ title: 'Attendance complete', body, url: '/' });
  const dead = [];
  await Promise.all(audience(fresh.subs, actor).map((s) => webpush.sendNotification(s.sub, payload, { TTL: 6 * 3600 }).catch((e) => { if (e && (e.statusCode === 404 || e.statusCode === 410)) dead.push(s); })));
  if (dead.length) { try { await G.clear(USERS_ID(), dead.map((s) => TAB + '!A' + s.row + ':D' + s.row)); } catch (e) {} }
  return true;
}
module.exports = { publicKey, subscribe, unsubscribe, status, pending, attendanceDone, _audience: audience };
