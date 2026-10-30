// Documents (admin only): Roles Clarity.
// POST { fn: 'roles.list' | 'roles.save' | 'roles.remove', args }
const G = require('./_lib/google');
const A = require('./_lib/auth');
const R = require('./_lib/roles');

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'POST only' });
  let body = req.body;
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch (e) { body = {}; } }
  body = body || {};

  const me = await A.guard(req, res); if (!me) return;
  if (!A.allow(me, res, 'admin')) return;                       // admins only, whatever permissions are ticked
  if (!R.ROUTES[body.fn]) return res.status(400).json({ ok: false, error: 'Bad function' });
  if (!G.enabled()) return res.status(500).json({ ok: false, error: 'Google connection is not set up' });
  try {
    const data = await R.ROUTES[body.fn](body.args || {}, me.name || me.email);
    return res.status(200).json({ ok: true, data });
  } catch (e) {
    return res.status(200).json({ ok: false, error: e.message });
  }
};
