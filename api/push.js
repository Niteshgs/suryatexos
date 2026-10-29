// Phone notification settings. POST { fn: 'push.status' | 'push.subscribe' | 'push.unsubscribe', args }
// Only Admin and Partner accounts can turn alerts on.
const G = require('./_lib/google');
const A = require('./_lib/auth');
const P = require('./_lib/push');

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'POST only' });
  let body = req.body;
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch (e) { body = {}; } }
  body = body || {};
  const me = await A.guard(req, res); if (!me) return;
  if (me.role !== 'Admin' && me.role !== 'Partner') return res.status(403).json({ ok: false, forbidden: true, error: 'Alerts are for Admin and Partner accounts' });
  if (!G.enabled()) return res.status(500).json({ ok: false, error: 'Google connection is not set up' });
  const a = body.args || {};
  try {
    let data;
    if (body.fn === 'push.status') data = await P.status(String(a.endpoint || ''));
    else if (body.fn === 'push.subscribe') data = await P.subscribe(me.email, me.role, a.sub);
    else if (body.fn === 'push.unsubscribe') data = await P.unsubscribe(String(a.endpoint || ''));
    else return res.status(400).json({ ok: false, error: 'Bad function' });
    return res.status(200).json({ ok: true, data });
  } catch (e) { return res.status(200).json({ ok: false, error: e.message }); }
};
