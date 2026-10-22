// Vercel serverless proxy → Apps Script Web App.
// Env: APPS_SCRIPT_URL, APPS_SCRIPT_TOKEN, APP_PASSCODE (optional)
const ACTIONS = ['getAll', 'saveStock', 'updateMin', 'addMaterial', 'addOrder', 'updateOrder'];

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'POST only' });

  let body = req.body;
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch (e) { body = {}; } }
  body = body || {};

  const pass = process.env.APP_PASSCODE;
  if (pass && body.passcode !== pass) return res.status(401).json({ ok: false, error: 'Wrong passcode' });
  if (!ACTIONS.includes(body.action)) return res.status(400).json({ ok: false, error: 'Bad action' });

  const url = process.env.APPS_SCRIPT_URL, token = process.env.APPS_SCRIPT_TOKEN;
  if (!url || !token) return res.status(500).json({ ok: false, error: 'Server not configured' });

  const { passcode, ...rest } = body;
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
