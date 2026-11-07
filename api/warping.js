// Warping: Planner (warping.plan) and Execute screen (warping.execute).
// POST { fn: 'warp.catalog' | 'warp.list' | 'warp.add' | 'warp.cancel' | 'warp.execList' | 'warp.setStatus', args }
const G = require('./_lib/google');
const A = require('./_lib/auth');
const W = require('./_lib/warping');

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'POST only' });
  let body = req.body;
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch (e) { body = {}; } }
  body = body || {};

  const me = await A.guard(req, res); if (!me) return;
  const boss = !!(me.isAdmin || me.role === 'Admin' || me.role === 'Partner');
  const need = A.ROUTE_PERMS.warping[body.fn];
  if (!need || !W.ROUTES[body.fn]) return res.status(400).json({ ok: false, error: 'Bad function' });
  if (!A.allow(me, res, need)) return;
  if (['warp.cancel', 'warp.toExecute', 'warp.archive', 'warp.unarchive', 'warp.archived', 'warp.analytics', 'warp.mnote'].includes(body.fn) && !boss) return res.status(403).json({ ok: false, error: 'Only Admin or Partner can do this' });
  if (!G.enabled()) return res.status(500).json({ ok: false, error: 'Google connection is not set up' });
  try {
    const data = await W.ROUTES[body.fn](body.args || {}, me.name || me.email);
    // management notes are only for Admin and Partner: strip them from everything else (a copy, so the shared cache is untouched)
    return res.status(200).json({ ok: true, data: boss || data === undefined ? data : JSON.parse(JSON.stringify(data, (k, v) => (k === 'mnote' ? undefined : v))) });
  } catch (e) {
    return res.status(200).json({ ok: false, error: e.message });
  }
};
