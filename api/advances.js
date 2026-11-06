// Advances, served straight from Google Sheets.
// POST { fn: 'advances.add' | 'advances.list', args }
// Callers must be signed in and have the matching permission (see _lib/auth.js).
const G = require('./_lib/google');
const A = require('./_lib/auth');
const V = require('./_lib/advances');

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'POST only' });
  let body = req.body;
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch (e) { body = {}; } }
  body = body || {};

  const me = await A.guard(req, res); if (!me) return;
  const need = A.ROUTE_PERMS.advances[body.fn];
  if (!need || !V.ROUTES[body.fn]) return res.status(400).json({ ok: false, error: 'Bad function' });
  if (!A.allow(me, res, need)) return;
  if (!G.enabled()) return res.status(500).json({ ok: false, error: 'Google connection is not set up' });

  try {
    const data = await V.ROUTES[body.fn](body.args || {}, me.name || me.email);
    return res.status(200).json({ ok: true, data });
  } catch (e) {
    return res.status(200).json({ ok: false, error: e.message });
  }
};
