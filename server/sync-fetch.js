'use strict';
/**
 * Pembantu UrlFetchApp.fetch: membaca permintaan (JSON) dari stdin, menjalankan fetch,
 * lalu menulis hasilnya (JSON) ke stdout. Dijalankan sebagai proses anak oleh gas-runtime.js.
 */
let input = '';
process.stdin.setEncoding('utf8');
process.stdin.on('data', (c) => { input += c; });
process.stdin.on('end', async () => {
  try {
    const req = JSON.parse(input);
    const headers = { ...(req.headers || {}) };
    if (req.contentType && !headers['Content-Type']) headers['Content-Type'] = req.contentType;
    const res = await fetch(req.url, {
      method: req.method || 'GET',
      headers,
      body: req.method === 'GET' || req.method === 'HEAD' ? undefined : req.payload,
      signal: AbortSignal.timeout(110000),
    });
    const body = await res.text();
    process.stdout.write(JSON.stringify({ status: res.status, body, headers: Object.fromEntries(res.headers) }));
  } catch (err) {
    process.stdout.write(JSON.stringify({ error: (err && (err.cause && err.cause.message || err.message)) || String(err) }));
  }
});
