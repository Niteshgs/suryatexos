// Vercel serverless proxy → Finished Goods Apps Script API (read-only).
// Env: FG_SCRIPT_URL, FG_SCRIPT_TOKEN, APP_PASSCODE (optional, shared with /api/sheet)
const ACTIONS = ['getAll', 'sweepPhotos'];

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
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
