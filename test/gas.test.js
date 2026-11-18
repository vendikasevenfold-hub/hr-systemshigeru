'use strict';
// Menjalankan fungsi asli Code.gs / Setup.gs di atas workbook memori (tanpa database).
process.env.TZ = 'Asia/Jakarta';
const test = require('node:test');
const assert = require('node:assert');
const { Workbook } = require('../server/workbook');
const { loadGas, formatDate, sessionOwner } = require('../server/gas-runtime');
const { encodeRows, decodeCell } = require('../server/store');

const wb = new Workbook();
const callable = loadGas(wb);
const g = globalThis;
['setupUsersTable', 'setupDatabase', 'setupPayrollModule', 'setupAbsensiModule', 'setupPermissionsModule'].forEach((f) => g[f]());

test('fungsi privat (berakhiran _) tidak bisa dipanggil dari browser', () => {
  assert.ok(callable.has('loginUser'));
  assert.ok(!callable.has('createSession_'));
  assert.ok(!callable.has('doGet'));
});

test('seed membuat sheet standar', () => {
  for (const n of ['Users', 'Karyawan', 'PermintaanLembur', 'MasterLokasiKerja', 'MasterLevelJabatan', 'MasterGajiKaryawan', 'PayrollInput', 'PayrollSettings', 'AbsensiHarian', 'Permissions']) {
    assert.ok(wb.getSheetByName(n), 'sheet ' + n);
  }
  assert.strictEqual(wb.getSheetByName('Sheet1'), null);
  assert.ok(wb.getSheetByName('Permissions').getLastRow() > 10);
});

test('login, sesi, dan manajemen user', () => {
  const bad = g.loginUser('admin', 'salah');
  assert.strictEqual(bad.success, false);
  const res = g.loginUser('admin', 'admin123');
  assert.strictEqual(res.success, true, res.message);
  assert.strictEqual(res.user.role, 'Admin');
  assert.strictEqual(res.user.permissions['users.view'], true);
  assert.strictEqual(g.validateSession(res.token, 'admin'), true);

  const add = g.saveNewUser({ username: 'mgr1', email: 'mgr1@x.co.id', password: 'p', role: 'Manager' }, res.token, 'admin');
  assert.strictEqual(add.success, true, add.message);
  const list = g.getUsersList(res.token, 'admin');
  assert.deepStrictEqual(list.map((u) => u.id), ['USR-001', 'USR-002']);
  assert.match(list[1].createdAt, /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/);
  assert.deepStrictEqual(g.getManagerList('MOD').map((m) => m.username), ['mgr1']);

  // Token dari loginUser dikenali server (dipakai untuk mewajibkan login di /api/run)
  assert.strictEqual(sessionOwner(res.token), 'admin');
  assert.strictEqual(sessionOwner('token-palsu'), null);
  assert.strictEqual(sessionOwner(undefined), null);

  g.logoutSession(res.token);
  assert.strictEqual(g.validateSession(res.token, 'admin'), false);
  assert.strictEqual(sessionOwner(res.token), null);
});

test('alur lembur: buat -> setujui manager -> setujui HRD', () => {
  const r = g.createLembur({ outlet: 'Head Office', periode: 'Jan 2026', idKaryawan: '000202333', namaKaryawan: 'Budi', jabatan: 'Cashier',
    managerApproval: 'mgr1', hari: 'Senin', tglLembur: '2026-01-05', jamMasuk: '09:00', jamKeluar: '20:00', keterangan: 'stok', createdBy: 'admin' });
  assert.strictEqual(r.success, true);
  let data = g.getLemburData('Admin', 'admin');
  assert.strictEqual(data.length, 1);
  assert.strictEqual(data[0].id, 'OVT-0001');
  assert.strictEqual(data[0].status, 'Pending');

  assert.strictEqual(g.updateOvertimeStatus('OVT-0001', 'Disetujui', 'mgr1').status, 'Pending HRD');
  assert.strictEqual(g.updateOvertimeStatus('OVT-0001', 'Disetujui', 'admin').status, 'Disetujui');
  data = g.getLemburData('Admin', 'admin');
  assert.match(data[0].catatan, /Disetujui Manager oleh mgr1 .* \| Disetujui HRD oleh admin/);

  // Realisasi dari ADMS: masuk 08:00, pulang 19:40 -> lebih 2 jam 40 menit -> 3 jam
  wb.getSheetByName('AbsensiHarian').appendRow(['000202333', 'Budi', 'MMJ-HOF', '2026-01-05', '08:00:00', '19:40:00', 'Hadir', 0, '', '', '']);
  data = g.getLemburData('Admin', 'admin');
  assert.strictEqual(data[0].totalJam, 3);
  assert.strictEqual(data[0].totalBiaya, 30000);

  assert.strictEqual(g.deleteLemburData('OVT-0001').success, true);
  assert.strictEqual(g.getLemburData('Admin', 'admin').length, 0);
});

test('master karyawan, mutasi, dan input harian', () => {
  assert.strictEqual(g.saveMasterGajiKaryawan({ idKaryawan: 'K001', nama: 'Sari', kodeLokasi: 'MMJ-GGP', jabatan: 'Cashier' }).success, true);
  const m = g.mutasiKaryawan({ idKaryawan: 'K001', kodeLokasiBaru: 'MMJ-HOF', alasan: 'tes' }, 'Admin', 'admin');
  assert.strictEqual(m.success, true, m.message);
  assert.strictEqual(g.getMasterGajiKaryawan()[0].kodeLokasi, 'MMJ-HOF');
  assert.strictEqual(g.getMutasiList('K001').length, 1);

  const s = g.saveInputHarian({ idKaryawan: 'K001', bulan: 2, tahun: 2026, inputBy: 'admin',
    hari: [{ tanggal: '2026-01-20', status: 'Hadir' }, { tanggal: '2026-01-21', status: 'Izin', terlambat: true }] });
  assert.strictEqual(s.success, true, s.message);
  const pin = g.getPayrollInput(2, 2026);
  assert.strictEqual(pin.length, 1);
  assert.strictEqual(pin[0].hariKerjaAktual, 1);
  assert.strictEqual(pin[0].jumlahIzin, 1);
  assert.strictEqual(pin[0].terlambat, 1);
});

test('lokasi: tambah, ubah kode, hapus', () => {
  assert.strictEqual(g.saveLokasiKerja({ kodeLokasi: 'MMJ-TES', namaLokasi: 'Tes', jadwalKerja: 'x' }).success, true);
  const u = g.saveLokasiKerja({ kodeLama: 'MMJ-TES', kodeLokasi: 'MMJ-TES2', namaLokasi: 'Tes 2', jadwalKerja: 'x' });
  assert.strictEqual(u.success, true, u.message);
  const tok = g.loginUser('admin', 'admin123').token;
  assert.strictEqual(g.deleteMasterLokasi('MMJ-TES2', tok, 'admin').success, true);
});

test('formatDate, nilai tanggal, dan encode/decode database', () => {
  const d = new Date(2026, 0, 5, 7, 3, 9);
  assert.strictEqual(formatDate(d, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss'), '2026-01-05 07:03:09');
  assert.strictEqual(formatDate(d, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm'), '05/01/2026 07:03');
  const json = encodeRows([['a', 1, true, d, '']]);
  const back = JSON.parse(json)[0].map(decodeCell);
  assert.strictEqual(back[3].getTime(), d.getTime());
  assert.ok(back[3] instanceof Date);
  const sh = wb.insertSheet('Tmp');
  sh.appendRow([d, new Date(1899, 11, 30, 8, 30), 5, false]);
  assert.deepStrictEqual(sh.getDataRange().getDisplayValues()[0], ['2026-01-05 07:03:09', '08:30', '5', 'FALSE']);
});
