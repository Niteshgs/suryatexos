// Sign-in endpoint for Suryatex OS.
// POST { action: 'login', credential } · 'me' · 'logout' · 'users' (admin) · 'saveUser' (admin)
const A = require('./_lib/auth');

function bodyOf(req) {
  let b = req.body;
  if (typeof b === 'string') { try { b = JSON.parse(b); } catch (e) { b = {}; } }
  return b || {};
}

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'POST only' });
  const body = bodyOf(req);

  try {
    if (body.action === 'login') {
      let g;
      try { g = await A.verifyGoogle(body.credential); }
      catch (e) { return res.status(401).json({ ok: false, error: e.code === 'BADTOKEN' ? e.message : 'Google sign-in failed. Please try again.' }); }
      const users = await A.loadUsers(true);
      const u = users.find((x) => x.email === g.email);
      if (!u) return res.status(403).json({ ok: false, error: g.email + ' is not set up for Suryatex OS yet. Ask the admin to add this email.' });
      if (!u.active) return res.status(403).json({ ok: false, error: 'Your access has been turned off. Please contact the admin.' });
      A.setSession(res, u.email);
      await A.touch(u);
      return res.status(200).json({ ok: true, me: A.meOf(u) });
    }

    if (body.action === 'logout') { A.clearSession(res); return res.status(200).json({ ok: true }); }

    if (body.action === 'me') {
      let me;
      try { me = await A.whoami(req); } catch (e) { return res.status(503).json({ ok: false, error: 'Could not check your access right now' }); }
      return res.status(200).json({ ok: true, me: me || null });
    }

    if (body.action === 'users' || body.action === 'saveUser') {
      const me = await A.guard(req, res); if (!me) return;
      if (!A.allow(me, res, 'admin')) return;
      if (body.action === 'users') return res.status(200).json({ ok: true, users: await A.listUsers(), catalog: A.CATALOG, roles: A.ROLES });
      try { return res.status(200).json({ ok: true, users: await A.saveUser(me, body.user || {}), catalog: A.CATALOG, roles: A.ROLES }); }
      catch (e) { return res.status(200).json({ ok: false, error: e.message }); }
    }

    return res.status(400).json({ ok: false, error: 'Bad action' });
  } catch (e) {
    return res.status(502).json({ ok: false, error: 'Sign-in error: ' + e.message });
  }
};
