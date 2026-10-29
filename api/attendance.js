// Staff Attendance and Leave Plan Form, served straight from Google Sheets.
// POST { fn: 'att.*' and 'leave.*' functions (see _lib/attendance.js), args }
// Callers must be signed in and have the matching permission (see _lib/auth.js).
const G = require('./_lib/google');
const A = require('./_lib/auth');
const E = require('./_lib/attendance');

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'POST only' });
  let body = req.body;
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch (e) { body = {}; } }
  body = body || {};

  // Public team-site leave form (team-suryatex.vercel.app): no sign-in, but only with the shared server-to-server token.
  const given = req.headers['x-team-token'];
  if (given !== undefined) {
    const want = process.env.TEAM_LEAVE_TOKEN || '';
    const h = (v) => require('crypto').createHash('sha256').update(String(v)).digest();
    if (!want || !require('crypto').timingSafeEqual(h(given), h(want))) return res.status(401).json({ ok: false, error: 'Not allowed' });
    if (!/^team\./.test(String(body.fn)) || !E.ROUTES[body.fn]) return res.status(400).json({ ok: false, error: 'Bad function' });
    if (!G.enabled()) return res.status(500).json({ ok: false, error: 'Google connection is not set up' });
    try { return res.status(200).json({ ok: true, data: await E.ROUTES[body.fn](body.args || {}, 'Team site (public form)') }); }
    catch (e) { return res.status(200).json({ ok: false, error: e.message }); }
  }

  const me = await A.guard(req, res); if (!me) return;
  const need = A.ROUTE_PERMS.attendance[body.fn];
  if (!need || !E.ROUTES[body.fn]) return res.status(400).json({ ok: false, error: 'Bad function' });
  if (!A.allow(me, res, need)) return;
  if (!G.enabled()) return res.status(500).json({ ok: false, error: 'Google connection is not set up' });

  // Attendance in-charge (markToday only): today's date, nothing else.
  const full = ['attendance.mark', 'attendance.view', 'attendance.edit', 'attendance.staff'].some((k) => me.access.includes(k));
  const inCharge = !full && me.access.includes('attendance.markToday');
  const today = new Date(Date.now() + 5.5 * 3600 * 1000).toISOString().slice(0, 10);   // India date
  if (inCharge && body.fn === 'att.save' && String((body.args || {}).date) !== today)
    return res.status(403).json({ ok: false, forbidden: true, error: 'You can only mark today\'s attendance' });

  try {
    let data = await E.ROUTES[body.fn](body.args || {}, (me.name || "") + " <" + me.email + "> · web", { email: me.email, role: me.role });
    if (inCharge && body.fn === 'att.getAll' && data && data.attendance) {
      const keep = {}; Object.keys(data.attendance).forEach((k) => { if (k.slice(-10) === today) keep[k] = data.attendance[k]; });
      data = Object.assign({}, data, { attendance: keep });
    }
    return res.status(200).json({ ok: true, data });
  } catch (e) {
    return res.status(200).json({ ok: false, error: e.message });
  }
};
