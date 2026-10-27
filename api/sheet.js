// Vercel serverless proxy → Apps Script Web App.
// Env: APPS_SCRIPT_URL, APPS_SCRIPT_TOKEN. Callers must be signed in and have the matching permission (see _lib/auth.js).
const A = require('./_lib/auth');

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'POST only' });

  let body = req.body;
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch (e) { body = {}; } }
  body = body || {};

  const me = await A.guard(req, res); if (!me) return;
  const need = A.ROUTE_PERMS.yarn[body.action];
  if (!need) return res.status(400).json({ ok: false, error: 'Bad action' });
  if (!A.allow(me, res, need)) return;

  const url = process.env.APPS_SCRIPT_URL, token = process.env.APPS_SCRIPT_TOKEN;
  if (!url || !token) return res.status(500).json({ ok: false, error: 'Server not configured' });

  const { passcode, ...rest } = body;   // (legacy field, never forwarded)
  try {
    const r = await fetch(url, {
      method: 'POST',
      redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ ...rest, token }),
    });
    const text = await r.text();
    let json;
    try { json = JSON.parse(text); } catch (e) {
      return res.status(502).json({ ok: false, error: 'Apps Script returned non-JSON (check deployment access = Anyone)' });
    }
    return res.status(200).json(json);
  } catch (e) {
    return res.status(502).json({ ok: false, error: 'Upstream error: ' + e.message });
  }
};
