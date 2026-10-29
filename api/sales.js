// Sales functions (Customer Touch Point Tracker), served straight from Google Sheets.
// POST { fn: 'touch.getAll' | 'touch.toggle' | 'touch.addClient' | 'touch.updateClient' | 'touch.archiveClient', args }
// Callers must be signed in and have the matching permission (see _lib/auth.js).
const G = require('./_lib/google');
const A = require('./_lib/auth');
const T = require('./_lib/touch');

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'POST only' });
  let body = req.body;
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch (e) { body = {}; } }
  body = body || {};

  const me = await A.guard(req, res); if (!me) return;
  const need = A.ROUTE_PERMS.sales[body.fn];
  if (!need || !T.ROUTES[body.fn]) return res.status(400).json({ ok: false, error: 'Bad function' });
  if (!A.allow(me, res, need)) return;
  if (!G.enabled()) return res.status(500).json({ ok: false, error: 'Google connection is not set up' });

  // Class A and Class B are separate permissions: check the class this call touches.
  const mayClass = (c) => me.access.includes('sales.touch' + c);
  const args = body.args || {};
  if (body.fn !== 'touch.getAll' && !mayClass(String(args.cls))) return res.status(403).json({ ok: false, forbidden: true, error: 'You do not have access to this' });

  try {
    let data = await T.ROUTES[body.fn](args, me.name || me.email);
    if (body.fn === 'touch.getAll') { data = Object.assign({}, data); ['A', 'B'].forEach((c) => { if (!mayClass(c)) delete data[c]; }); }
    return res.status(200).json({ ok: true, data });
  } catch (e) {
    return res.status(200).json({ ok: false, error: e.message });
  }
};
