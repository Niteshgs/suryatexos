// Vercel serverless proxy → Suryatex OS Hub (standalone Apps Script).
// Env: HUB_SCRIPT_URL, HUB_SCRIPT_TOKEN, APP_PASSCODE (optional, shared with the other routes)
const FUNCTIONS = [
  'knitted.getAll', 'knitted.saveBeams', 'knitted.verifyPin',
  'knitted.addTape', 'knitted.removeTape', 'knitted.renameTape',
];

async function callHub(fn, args, meta) {
  const url = process.env.HUB_SCRIPT_URL, token = process.env.HUB_SCRIPT_TOKEN;
  if (!url || !token) return { status: 500, json: { ok: false, error: 'This function is not connected yet' } };
  const r = await fetch(url, {
    method: 'POST', redirect: 'follow',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ token, fn, args: args || {}, meta: meta || {} }),
  });
  const text = await r.text();
  try { return { status: 200, json: JSON.parse(text) }; }
  catch (e) { return { status: 502, json: { ok: false, error: 'Hub returned non-JSON (check deployment access = Anyone)' } }; }
}

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');

  // Preview builds only: GET returns counts (no sheet data) so the connection can be checked.
  if (req.method === 'GET' && process.env.VERCEL_ENV === 'preview') {
    try {
      const t0 = Date.now();
      const r = await callHub('knitted.getAll');
      const d = r.json.data;
      return res.status(200).json({ ok: r.json.ok, error: r.json.error || '', ms: Date.now() - t0,
        knitted: d ? { machines: Object.keys(d.cycleData).length, tapes: d.tapeList.length,
          beamsWithTape: Object.values(d.cycleData).flat().filter((b) => b.tape).length } : null });
    } catch (e) { return res.status(502).json({ ok: false, error: e.message }); }
  }
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'POST only' });

  let body = req.body;
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch (e) { body = {}; } }
  body = body || {};

  const pass = process.env.APP_PASSCODE;
  if (pass && body.passcode !== pass) return res.status(401).json({ ok: false, error: 'Wrong passcode' });
  if (!FUNCTIONS.includes(body.fn)) return res.status(400).json({ ok: false, error: 'Bad function' });

  const m = body.meta || {};
  const meta = { user: String(m.user || '').slice(0, 60), device: String(m.device || '').slice(0, 80), pin: String(m.pin || '').slice(0, 20) };
  try {
    const r = await callHub(body.fn, body.args, meta);
    return res.status(r.status).json(r.json);
  } catch (e) {
    return res.status(502).json({ ok: false, error: 'Upstream error: ' + e.message });
  }
};
