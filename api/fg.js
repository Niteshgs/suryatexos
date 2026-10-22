// Vercel serverless proxy → Finished Goods Apps Script API (read-only).
// Env: FG_SCRIPT_URL, FG_SCRIPT_TOKEN, APP_PASSCODE (optional, shared with /api/sheet)
const ACTIONS = ['getAll', 'sweepPhotos'];

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  // Preview builds only: GET returns counts (no sheet data) so the connection can be checked.
  if (req.method === 'GET' && process.env.VERCEL_ENV === 'preview') return healthCheck(res);
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'POST only' });

  let body = req.body;
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch (e) { body = {}; } }
  body = body || {};

  const pass = process.env.APP_PASSCODE;
  if (pass && body.passcode !== pass) return res.status(401).json({ ok: false, error: 'Wrong passcode' });
  if (!ACTIONS.includes(body.action)) return res.status(400).json({ ok: false, error: 'Bad action' });

  const url = process.env.FG_SCRIPT_URL, token = process.env.FG_SCRIPT_TOKEN;
  if (!url || !token) return res.status(500).json({ ok: false, error: 'Finished Goods is not connected yet' });

  const payload = { action: body.action, token };
  if (body.action === 'sweepPhotos') {
    payload.force = !!body.force;
    if (Array.isArray(body.rows)) payload.rows = body.rows.slice(0, 500).map(String);
  }
  try {
    const r = await fetch(url, {
      method: 'POST',
      redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
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

async function healthCheck(res) {
  const url = process.env.FG_SCRIPT_URL, token = process.env.FG_SCRIPT_TOKEN;
  if (!url || !token) return res.status(500).json({ ok: false, error: 'not configured' });
  const post = async (payload) => {
    const t0 = Date.now();
    const r = await fetch(url, { method: 'POST', redirect: 'follow', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ ...payload, token }) });
    const text = await r.text();
    let j; try { j = JSON.parse(text); } catch (e) { return { ok: false, error: 'non-JSON: ' + text.slice(0, 120), ms: Date.now() - t0 }; }
    j.ms = Date.now() - t0; j.bytes = text.length; return j;
  };
  try {
    const a = await post({ action: 'getAll' });
    const out = { getAll: { ok: a.ok, error: a.error || '', ms: a.ms, bytes: a.bytes } };
    if (a.ok) {
      const d = a.data;
      out.getAll.fgRows = d.fg.rows.length; out.getAll.fgError = d.fg.error; out.getAll.fgUpdated = d.fg.updated;
      out.getAll.pos = d.po.pos.length; out.getAll.poError = d.po.error;
      out.getAll.sweepRows = d.sweep.rows.length; out.getAll.sweepError = d.sweep.error;
      out.getAll.machines = [...new Set(d.fg.rows.map((r) => r.machine))];
    }
    const p = await post({ action: 'sweepPhotos' });
    out.photos = { ok: p.ok, error: p.error || (p.data && p.data.error) || '', ms: p.ms, bytes: p.bytes,
      rowsWithPhotos: p.data ? Object.keys(p.data.images).length : 0, more: p.data ? p.data.more.length : 0, skipped: p.data ? p.data.skipped : 0 };
    return res.status(200).json(out);
  } catch (e) { return res.status(502).json({ ok: false, error: e.message }); }
}
