// Direct Google Sheets access for Suryatex OS (no Apps Script).
// Uses the "suryatex-os" service account. Env: GOOGLE_SERVICE_KEY = the service account's JSON key.
// Only sheets that have been shared with the service account's email can be reached.
const crypto = require('crypto');

const SCOPES = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.readonly';
let cached = { token: null, exp: 0 };

function serviceKey() {
  const raw = process.env.GOOGLE_SERVICE_KEY;
  if (!raw) return null;
  try { const k = JSON.parse(raw); return k.client_email && k.private_key ? k : null; } catch (e) { return null; }
}
function enabled() { return !!serviceKey(); }

const b64url = (b) => Buffer.from(b).toString('base64').replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_');

async function token() {
  if (cached.token && Date.now() < cached.exp - 60000) return cached.token;
  const key = serviceKey();
  if (!key) throw new Error('Google connection is not set up');
  const now = Math.floor(Date.now() / 1000);
  const head = b64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const claim = b64url(JSON.stringify({ iss: key.client_email, scope: SCOPES, aud: 'https://oauth2.googleapis.com/token', iat: now, exp: now + 3600 }));
  const sig = crypto.createSign('RSA-SHA256').update(head + '.' + claim).sign(key.private_key);
  const r = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: 'grant_type=' + encodeURIComponent('urn:ietf:params:oauth:grant-type:jwt-bearer') + '&assertion=' + head + '.' + claim + '.' + b64url(sig),
  });
  const j = await r.json();
  if (!j.access_token) throw new Error('Google sign-in failed: ' + (j.error_description || j.error || r.status));
  cached = { token: j.access_token, exp: Date.now() + (j.expires_in || 3600) * 1000 };
  return cached.token;
}

async function api(method, url, body) {
  const r = await fetch(url, {
    method, headers: { Authorization: 'Bearer ' + (await token()), 'Content-Type': 'application/json' },
    body: body ? JSON.stringify(body) : undefined,
  });
  const text = await r.text();
  let j = {}; try { j = text ? JSON.parse(text) : {}; } catch (e) {}
  if (!r.ok) {
    const msg = (j.error && j.error.message) || text.slice(0, 200);
    if (r.status === 403 || r.status === 404) throw new Error('Sheet not reachable — share it with ' + ((serviceKey() || {}).client_email || 'the app account') + ' as Editor (' + msg + ')');
    throw new Error('Google Sheets error ' + r.status + ': ' + msg);
  }
  return j;
}

const base = (id) => 'https://sheets.googleapis.com/v4/spreadsheets/' + encodeURIComponent(id);
const q = (s) => "'" + String(s).replace(/'/g, "''") + "'";   // quote a tab name for A1 ranges

// values like Apps Script getValues(): numbers stay numbers, empty cells are ''
async function batchGet(id, ranges, dates, formatted) {   // formatted=true returns cells as they look on screen ("40%", "Rs. 1,500")
  const p = ranges.map((r) => 'ranges=' + encodeURIComponent(r)).join('&');
  const j = await api('GET', base(id) + '/values:batchGet?' + p + '&valueRenderOption=' + (formatted ? 'FORMATTED_VALUE' : 'UNFORMATTED_VALUE') + '&dateTimeRenderOption=' + (dates === 'serial' ? 'SERIAL_NUMBER' : 'FORMATTED_STRING') + '&majorDimension=ROWS');
  return (j.valueRanges || []).map((v) => v.values || []);
}
async function batchUpdate(id, data, inputOption) {
  if (!data.length) return;
  await api('POST', base(id) + '/values:batchUpdate', { valueInputOption: inputOption || 'RAW', data });
}
async function append(id, range, rows, inputOption) {
  await api('POST', base(id) + '/values/' + encodeURIComponent(range) + ':append?valueInputOption=' + (inputOption || 'RAW') + '&insertDataOption=INSERT_ROWS', { values: rows });
}
async function sheetProps(id) {
  const j = await api('GET', base(id) + '?fields=sheets.properties(title,sheetId,gridProperties(frozenRowCount,columnCount))');
  return (j.sheets || []).map((s) => s.properties);
}
async function format(id, requests) {
  if (requests.length) await api('POST', base(id) + ':batchUpdate', { requests });
}
async function sheetTitles(id) {
  const j = await api('GET', base(id) + '?fields=sheets.properties.title');
  return (j.sheets || []).map((s) => s.properties.title);
}

// read-only: can the service account edit this file, and which ranges are protected?
async function accessReport(id) {
  const out = {};
  try {
    const d = await api('GET', 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(id) + '?fields=name,capabilities(canEdit),owners(emailAddress)&supportsAllDrives=true');
    out.file = d.name; out.canEdit = d.capabilities && d.capabilities.canEdit;
  } catch (e) { out.driveError = e.message; }
  try {
    const j = await api('GET', base(id) + '?fields=sheets(properties(title,sheetId),protectedRanges(range,warningOnly,description,editors(users)))');
    out.protected = [];
    (j.sheets || []).forEach((sh) => (sh.protectedRanges || []).forEach((pr) => out.protected.push({
      tab: sh.properties.title, wholeTab: !pr.range || (pr.range.startRowIndex === undefined && pr.range.startColumnIndex === undefined),
      range: pr.range, warningOnly: !!pr.warningOnly, editors: pr.editors && pr.editors.users ? pr.editors.users.length : 0,
      appIsEditor: !!(pr.editors && pr.editors.users && pr.editors.users.includes((serviceKey() || {}).client_email)) })));
  } catch (e) { out.sheetsError = e.message; }
  return out;
}

module.exports = { accessReport, enabled, token, batchGet, batchUpdate, append, sheetTitles, sheetProps, format, q, serviceEmail: () => (serviceKey() || {}).client_email };
