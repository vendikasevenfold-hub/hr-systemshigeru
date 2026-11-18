'use strict';
/**
 * Server Shigeru HR.
 *  - GET  /                 -> public/index.html (UI yang sama dengan versi Apps Script)
 *  - POST /api/run/:fungsi  -> menjalankan fungsi di Code.gs / Setup.gs (pengganti google.script.run)
 *  - GET  /healthz          -> cek kesehatan untuk Docker
 */
const path = require('path');
const express = require('express');
const { Workbook } = require('./workbook');
const { loadGas, sessionOwner } = require('./gas-runtime');
const { createStore, encodeChanges } = require('./store');

const PORT = Number(process.env.PORT) || 3000;
// Di Docker: 0.0.0.0. Di VPS tanpa Docker: 127.0.0.1, supaya hanya Apache yang bisa menghubungi.
const HOST = process.env.HOST || '0.0.0.0';

// Fungsi yang boleh dipanggil SEBELUM login. Semua fungsi lain wajib membawa token sesi
// yang sah (header X-HR-Token), supaya data karyawan/gaji tidak bisa diambil tanpa login.
const PUBLIC_FUNCTIONS = new Set(['loginUser', 'validateSession', 'logoutSession']);
const DATABASE_URL = process.env.DATABASE_URL;
const SEED_ON_EMPTY = String(process.env.SEED_ON_EMPTY || 'true').toLowerCase() === 'true';

// Fungsi setup untuk database yang masih kosong. Urutan penting: setupUsersTable dulu
// supaya sheet Users memakai format 6 kolom (setupDatabase membuat format lama 4 kolom).
const SEED_FUNCTIONS = ['setupUsersTable', 'setupDatabase', 'setupPayrollModule', 'setupAbsensiModule', 'setupPermissionsModule'];

async function main() {
  if (!DATABASE_URL) throw new Error('DATABASE_URL belum diisi.');

  // Tanggal/jam di Code.gs dihitung dengan zona waktu proses. Pastikan sama dengan Asia/Jakarta.
  const tzProses = Intl.DateTimeFormat().resolvedOptions().timeZone;
  const tzScript = process.env.TZ || 'Asia/Jakarta';
  if (tzProses !== tzScript) {
    console.warn(`PERINGATAN: zona waktu proses ${tzProses}, seharusnya ${tzScript}. ` +
      'Set TZ=Asia/Jakarta di .env atau ubah zona waktu server ke (UTC+07:00) Bangkok, Hanoi, Jakarta.');
  }

  const store = createStore(DATABASE_URL);
  await store.init();

  const workbook = new Workbook();
  const jumlahSheet = await store.loadInto(workbook);
  const callable = loadGas(workbook);

  // Simpan perubahan secara berurutan (satu antrean), supaya urutan tulis ke database sama
  // dengan urutan panggilan.
  let saveChain = Promise.resolve();
  const persist = () => {
    const encoded = encodeChanges(workbook.takeChanges());
    const p = saveChain.then(() => store.save(encoded));
    saveChain = p.catch(() => {});
    return p;
  };

  if (jumlahSheet === 0 && SEED_ON_EMPTY) {
    for (const fn of SEED_FUNCTIONS) globalThis[fn]();
    await persist();
    console.log('Database kosong: sheet standar dan user admin bawaan dibuat (admin / admin123). Segera ganti password-nya.');
  }
  console.log(`Workbook dimuat: ${workbook.getSheets().length} sheet. Fungsi yang bisa dipanggil: ${callable.size}.`);

  const app = express();
  app.disable('x-powered-by');
  app.use(express.json({ limit: '25mb' }));

  app.get('/healthz', async (_req, res) => {
    try { await store.pool.query('SELECT 1'); res.json({ ok: true }); } catch (e) { res.status(503).json({ ok: false }); }
  });

  app.post('/api/run/:fn', async (req, res) => {
    const name = req.params.fn;
    if (!callable.has(name)) {
      return res.status(404).json({ ok: false, error: 'Fungsi tidak ditemukan: ' + name });
    }
    if (!PUBLIC_FUNCTIONS.has(name) && !sessionOwner(req.get('X-HR-Token'))) {
      return res.status(401).json({ ok: false, sessionExpired: true, error: 'Sesi login tidak valid atau sudah berakhir. Silakan login ulang.' });
    }
    const args = Array.isArray(req.body && req.body.args) ? req.body.args : [];
    let result, error;
    try {
      result = globalThis[name](...args);
    } catch (err) {
      error = err;
    }
    try {
      // Sama seperti Apps Script: perubahan sebelum error tetap tersimpan.
      await persist();
    } catch (err) {
      console.error('Gagal menyimpan ke database setelah', name, err);
      return res.status(500).json({ ok: false, error: 'Gagal menyimpan ke database: ' + err.message });
    }
    if (error) {
      console.error('Error di', name + ':', error && error.stack || error);
      return res.status(500).json({ ok: false, error: (error && error.message) || String(error) });
    }
    res.json({ ok: true, result: result === undefined ? null : result });
  });

  app.use(express.static(path.join(__dirname, '..', 'public'), { index: 'index.html', maxAge: 0 }));

  const server = app.listen(PORT, HOST, () => console.log(`Shigeru HR berjalan di ${HOST}:${PORT}.`));

  const stop = async () => {
    server.close();
    try { await saveChain; await store.pool.end(); } finally { process.exit(0); }
  };
  process.on('SIGTERM', stop);
  process.on('SIGINT', stop);
}

main().catch((err) => {
  console.error('Gagal menjalankan server:', err);
  process.exit(1);
});
