/**
 * SuryatexOS_API.gs — read-only data feed for Suryatex OS (suryatexos.vercel.app)
 *
 * Add this as a NEW file in the existing "Snapweaves Stock" Apps Script project.
 * It does not change Code.gs or Index.html. It reuses their readers (getAllData, getSweepImages)
 * and only adds doPost, which the current app does not use.
 *
 * Setup (once):
 *   1. In the Apps Script editor: Files "+" -> Script -> name it SuryatexOS_API -> paste this file.
 *   2. Deploy -> New deployment -> type: Web app -> Execute as: Me, Who has access: Anyone -> Deploy.
 *      This creates a SECOND deployment with its own URL. Your Google Sites deployment keeps
 *      running the version it already has, so nothing there changes.
 *   3. Send the new Web app URL (ends in /exec) to be added to Suryatex OS.
 *
 * Safe for the existing app: it only reads the sheet and answers only requests carrying SXOS_TOKEN.
 */

var SXOS_TOKEN = '6J8-WbbTJ9uGnjIcHP3km_pODGT0pSC8';   // must match FG_SCRIPT_TOKEN on Vercel
var SXOS_PHOTO_PAGE_CHARS = 3000000;                    // keeps each reply under Vercel's 4.5 MB limit

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (x) {}
  if (body.token !== SXOS_TOKEN) return sxosJson_({ ok: false, error: 'Unauthorized' });
  try {
    if (body.action === 'getAll') return sxosJson_({ ok: true, data: getAllData() });
    if (body.action === 'sweepPhotos') return sxosJson_({ ok: true, data: sxosPhotos_(!!body.force, body.rows || null) });
    return sxosJson_({ ok: false, error: 'Bad action' });
  } catch (err) {
    return sxosJson_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

// Optional check: pick "sxosTest" in the function menu and press Run. It only reads, and logs what it found.
function sxosTest() {
  var d = getAllData();
  Logger.log('F.Goods rows: ' + d.fg.rows.length + (d.fg.error ? '  ERROR: ' + d.fg.error : ''));
  Logger.log('POs: ' + d.po.pos.length + (d.po.error ? '  ERROR: ' + d.po.error : ''));
  Logger.log('Sweep rows: ' + d.sweep.rows.length + (d.sweep.error ? '  ERROR: ' + d.sweep.error : ''));
  var p = sxosPhotos_(false, null), n = 0;
  for (var k in p.images) n += p.images[k].length;
  Logger.log('Photos in first page: ' + n + ', rows in later pages: ' + p.more.length + (p.error ? '  ERROR: ' + p.error : ''));
}

function sxosJson_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* Stock SWEEP photos in pages: { images: { "<sheet row>": [dataUri] }, more: [rows left], skipped, error }.
 * First call with rows = null; then call again with the "more" list until it is empty. */
function sxosPhotos_(force, rows) {
  var all = getSweepImages(force && !rows);               // from Code.gs (cached for 6 hours)
  var out = { images: {}, more: [], skipped: all.skipped || 0, error: all.error || '' };
  var keys = rows ? rows.map(String) : Object.keys(all.images || {});
  var used = 0;
  for (var i = 0; i < keys.length; i++) {
    var imgs = (all.images || {})[keys[i]];
    if (!imgs) continue;
    var size = 0;
    for (var j = 0; j < imgs.length; j++) size += imgs[j].length;
    if (size > SXOS_PHOTO_PAGE_CHARS * 1.4) { out.skipped++; continue; }   // too big to send at all
    if (used && used + size > SXOS_PHOTO_PAGE_CHARS) { out.more.push(keys[i]); continue; }
    out.images[keys[i]] = imgs;
    used += size;
  }
  return out;
}
