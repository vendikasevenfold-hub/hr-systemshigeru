/**
 * Code.gs - Server-Side Backend Shigeru HR PT Mideli Makmur Jaya
 */

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Shigeru HR - PT Mideli Makmur Jaya')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// -------------------------------------------------------------
// AUTENTIKASI & MASTER DATA
// -------------------------------------------------------------
/**
 * Integrasi Login Dinamis dari Sheet Users (Standar 6 Kolom)
 */
function loginUser(username, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Users");

    if (!sheet) {
      return { success: false, message: "Sheet Users tidak ditemukan!" };
    }

    const data = sheet.getDataRange().getValues();

    // Looping data user (dimulai dari baris ke-2 / index 1)
    for (let i = 1; i < data.length; i++) {
      const dbId = data[i][0] ? data[i][0].toString().trim() : "";       // Kolom A (ID)
      const dbUsername = data[i][1] ? data[i][1].toString().trim() : ""; // Kolom B (Username)
      const dbEmail = data[i][2] ? data[i][2].toString().trim() : "";    // Kolom C (Email)
      const dbPassword = data[i][3] ? data[i][3].toString().trim() : ""; // Kolom D (Password)
      const dbRole = data[i][4] ? data[i][4].toString().trim() : "";     // Kolom E (Role)

      if (dbUsername.toLowerCase() === username.toLowerCase() && dbPassword === password) {
        return {
          success: true,
          token: createSession_(dbUsername),
          user: {
            id: dbId,
            username: dbUsername,
            email: dbEmail,
            role: dbRole,
            permissions: computeEffectivePermissions_(dbId, dbUsername, dbRole),
            lokasiAkses: computeEffectiveLokasi_(dbId, dbUsername, dbRole)
          }
        };
      }
    }

    return { success: false, message: "Username atau Password salah!" };
  } catch (error) {
    return { success: false, message: "Error Login: " + error.message };
  }
}

// ---------------- SESI LOGIN (token server, kedaluwarsa otomatis) ----------------
// Token dibuat saat login dan disimpan di CacheService (maks 6 jam, diperpanjang setiap dipakai).
// Link yang dibuka di browser/tab lain tidak membawa token -> selalu masuk portal login dulu.
const SESSION_TTL_SEC_ = 6 * 60 * 60;

function createSession_(username) {
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put("sess_" + token, String(username), SESSION_TTL_SEC_);
  return token;
}

function validateSession(token, username) {
  try {
    if (!token || !username) return false;
    const cache = CacheService.getScriptCache();
    const owner = cache.get("sess_" + token);
    if (!owner || String(owner).toLowerCase() !== String(username).toLowerCase()) return false;
    cache.put("sess_" + token, owner, SESSION_TTL_SEC_); // perpanjang
    return true;
  } catch (e) { return false; }
}

/** Pastikan pemanggil punya sesi login yang sah DAN berperan Admin/HR; lempar error bila tidak. Kembalikan data user. */
function requireHrSession_(token, username) {
  if (!validateSession(token, username)) throw new Error("Sesi login tidak valid atau sudah berakhir. Silakan login ulang.");
  const u = findUserByUsername_(username);
  if (!u || (u.role !== 'Admin' && u.role !== 'HR')) throw new Error("Akses ditolak. Hanya Admin dan HR yang dapat mengelola user.");
  return u;
}

function logoutSession(token) {
  try { if (token) CacheService.getScriptCache().remove("sess_" + token); } catch (e) { /* abaikan */ }
  return true;
}

// =================================================================
// MODUL HAK AKSES (PERMISSIONS) - per Role & per User
// =================================================================

// Daftar kanonik seluruh modul & fitur granular di dalam aplikasi.
// "view" adalah akses dasar untuk melihat/membuka modul (tab/menu sidebar).
// Fitur lain (di luar "view") adalah tombol/aksi spesifik di dalam modul tsb.
const MODULE_DEFINITIONS_ = [
  { key: 'master-karyawan', label: 'Kepegawaian - Data Master Karyawan',  hrOnly: true,  features: ['tambah_lokasi', 'tambah_karyawan', 'hapus_lokasi', 'hapus_karyawan'] },
  { key: 'lembur',          label: 'Lembur - Form & Data Lembur',         hrOnly: false, features: ['form', 'export'] },
  { key: 'input-bulanan',   label: 'Input Data Bulanan',                       hrOnly: true,  features: ['sync', 'export'] },
  { key: 'absensi',         label: 'Absensi (ADMS)',                          hrOnly: true,  features: ['sync', 'tarik'] },
  { key: 'users',           label: 'Pengaturan - Manajemen User',         hrOnly: true,  features: ['manage_permissions'] }
];

const PERMISSION_FEATURE_LABELS_ = {
  view: 'Lihat / Buka Modul',
  manage_permissions: 'Atur Hak Akses',
  form: 'Buat / Isi Form Lembur',
  export: 'Export (CSV / Excel)',
  sync: 'Sync dari ADMS',
  tarik: 'Tarik ke Input Data Bulanan',
  tambah_lokasi: 'Tambah / Edit Lokasi',
  tambah_karyawan: 'Tambah Karyawan per Lokasi',
  hapus_lokasi: 'Hapus Lokasi',
  hapus_karyawan: 'Hapus Karyawan'
};

// Default akses per Role, dipakai sebagai baseline saat sheet Permissions
// kosong/baru dibuat. Nilai ini meniru perilaku hak akses yang SUDAH berjalan
// di aplikasi sebelum fitur ini ada (tab-hr-only, dsb), supaya tidak ada
// perubahan perilaku mendadak begitu fitur ini aktif.
function defaultAllowedForRole_(role, moduleKey, featureKey) {
  if (role === 'Admin') return true; // Admin selalu penuh, anti terkunci sendiri

  const mod = MODULE_DEFINITIONS_.find(function (m) { return m.key === moduleKey; });
  if (!mod) return false;

  if (role === 'HR') {
    return true; // HR = setara Admin untuk semua modul (perilaku lama)
  }

  // Manager & MOD
  if (mod.hrOnly) return false; // modul yang dulu tab-hr-only tertutup untuk Manager/MOD
  if (moduleKey === 'lembur' && featureKey === 'export') {
    return role !== 'MOD'; // MOD dulu disembunyikan tombol export lembur-nya
  }
  return true;
}

function getPermissionsMatrix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Permissions");
  const rows = [];
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[0]) continue;
      rows.push({
        id: r[0],
        scopeType: r[1],
        scopeValue: r[2],
        moduleKey: r[3],
        featureKey: r[4],
        allowed: (r[5] === true || r[5] === 'TRUE')
      });
    }
  }
  return {
    modules: MODULE_DEFINITIONS_,
    featureLabels: PERMISSION_FEATURE_LABELS_,
    roles: ['Admin', 'HR', 'Manager', 'MOD'],
    users: loadUsersInternal_().map(function (u) {
      return { id: u.id, username: u.username, role: u.role };
    }),
    rows: rows,
    lokasi: getLokasiKerjaList().map(function (l) { return { kodeLokasi: l.kodeLokasi, namaLokasi: l.namaLokasi }; })
  };
}

function savePermissionsMatrix(payload) {
  if (!validateSession(payload && payload.token, payload && payload.requesterUsername)) {
    throw new Error("Sesi login tidak valid atau sudah berakhir. Silakan login ulang.");
  }
  const requester = findUserByUsername_(payload && payload.requesterUsername);
  if (!requester || requester.role !== 'Admin') {
    throw new Error("Akses ditolak. Hanya Admin yang dapat mengubah hak akses.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Permissions");
  if (!sheet) {
    sheet = ss.insertSheet("Permissions");
    sheet.appendRow(["ID", "ScopeType", "ScopeValue", "ModuleKey", "FeatureKey", "Allowed"]);
    sheet.setFrozenRows(1);
  }

  // Hapus dulu aturan lokasi lama untuk scope yang sedang disimpan (Role / User), lalu tulis ulang dari "changes".
  const clearScopes = (payload && payload.clearLokasiScopes) || [];
  if (clearScopes.length) {
    const cur = sheet.getDataRange().getValues();
    for (let i = cur.length - 1; i >= 1; i--) {
      if (cur[i][3] !== 'lokasi') continue;
      const hit = clearScopes.some(function (s) { return s.scopeType === cur[i][1] && s.scopeValue === cur[i][2]; });
      if (hit) sheet.deleteRow(i + 1);
    }
  }

  const data = sheet.getDataRange().getValues();
  const rowIndexByKey = {};
  let maxIdNum = 0;
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const key = [r[1], r[2], r[3], r[4]].join("|");
    rowIndexByKey[key] = i + 1; // nomor baris sheet (1-based)
    const m = String(r[0]).match(/PERM-(\d+)/);
    if (m) maxIdNum = Math.max(maxIdNum, parseInt(m[1], 10));
  }

  const changes = (payload && payload.changes) || [];
  changes.forEach(function (c) {
    const key = [c.scopeType, c.scopeValue, c.moduleKey, c.featureKey].join("|");
    if (rowIndexByKey[key]) {
      sheet.getRange(rowIndexByKey[key], 6).setValue(!!c.allowed);
    } else {
      maxIdNum++;
      const newId = "PERM-" + String(maxIdNum).padStart(4, "0");
      sheet.appendRow([newId, c.scopeType, c.scopeValue, c.moduleKey, c.featureKey, !!c.allowed]);
      rowIndexByKey[key] = sheet.getLastRow();
    }
  });

  return { success: true, message: "Hak akses berhasil disimpan." };
}

function findUserByUsername_(username) {
  if (!username) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().trim().toLowerCase() === username.toString().trim().toLowerCase()) {
      return { id: data[i][0], username: data[i][1], email: data[i][2], role: data[i][4] };
    }
  }
  return null;
}

// Menghitung hak akses efektif untuk satu user: default Role -> ditimpa
// override Role (jika ada di sheet Permissions) -> ditimpa override per-User
// (prioritas tertinggi, kecuali untuk Admin supaya Admin tidak bisa terkunci).
function computeEffectivePermissions_(userId, username, role) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Permissions");
  const rows = [];
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[0]) continue;
      rows.push({
        scopeType: r[1], scopeValue: r[2], moduleKey: r[3], featureKey: r[4],
        allowed: (r[5] === true || r[5] === 'TRUE')
      });
    }
  }

  const result = {};
  MODULE_DEFINITIONS_.forEach(function (mod) {
    const featureKeys = ['view'].concat(mod.features || []);
    featureKeys.forEach(function (feat) {
      let allowed = defaultAllowedForRole_(role, mod.key, feat);

      rows.forEach(function (r) {
        if (r.scopeType === 'ROLE' && r.scopeValue === role && r.moduleKey === mod.key && r.featureKey === feat) {
          allowed = r.allowed;
        }
      });

      if (role !== 'Admin') {
        rows.forEach(function (r) {
          if (r.scopeType === 'USER' && (r.scopeValue === userId || r.scopeValue === username) && r.moduleKey === mod.key && r.featureKey === feat) {
            allowed = r.allowed;
          }
        });
      }

      result[mod.key + '.' + feat] = allowed;
    });
  });
  return result;
}

// Lokasi yang boleh diakses user. null = semua lokasi.
// Disimpan di sheet Permissions dengan ModuleKey "lokasi": FeatureKey = kode lokasi (atau "*" = semua lokasi).
// Prioritas: aturan per-User (jika ada) menggantikan aturan Role.
// Admin dan HR tanpa aturan = semua lokasi. Role lain (Manager, MOD, dst) tanpa aturan = TIDAK ADA lokasi
// (lokasi harus dibuka eksplisit oleh Admin di Manajemen User > Akses Lokasi).
function computeEffectiveLokasi_(userId, username, role) {
  if (role === 'Admin') return null;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Permissions");
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const userRows = [], roleRows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0] || r[3] !== 'lokasi') continue;
    const row = { featureKey: String(r[4]).trim(), allowed: (r[5] === true || r[5] === 'TRUE') };
    if (r[1] === 'USER' && (r[2] === userId || r[2] === username)) userRows.push(row);
    else if (r[1] === 'ROLE' && r[2] === role) roleRows.push(row);
  }
  const rows = userRows.length ? userRows : roleRows;
  if (!rows.length) return (role === 'HR') ? null : [];
  if (rows.some(function (x) { return x.featureKey === '*' && x.allowed; })) return null;
  return rows.filter(function (x) { return x.featureKey !== '*' && x.allowed; }).map(function (x) { return x.featureKey; });
}

// Dipakai klien: hak akses modul/fitur + daftar lokasi yang boleh diakses.
function getMyAccess(username) {
  const user = findUserByUsername_(username);
  if (!user) throw new Error("User tidak ditemukan.");
  return {
    permissions: computeEffectivePermissions_(user.id, user.username, user.role),
    lokasiAkses: computeEffectiveLokasi_(user.id, user.username, user.role)
  };
}

// Untuk filter di sisi server: null = semua lokasi, selain itu daftar kode & nama lokasi (huruf kecil).
function getAllowedLokasiForUser_(username) {
  const user = findUserByUsername_(username);
  if (!user) return { kodes: [], names: [] };
  const kodes = computeEffectiveLokasi_(user.id, user.username, user.role);
  if (kodes === null) return null;
  const set = {}; kodes.forEach(function (k) { set[k.toLowerCase()] = true; });
  const names = [];
  getLokasiKerjaList().forEach(function (l) {
    if (set[String(l.kodeLokasi).trim().toLowerCase()]) names.push(String(l.namaLokasi).trim().toLowerCase());
  });
  return { kodes: Object.keys(set), names: names };
}



function getMasterData() {
  const locations = [];
  try {
    const seen = {};
    getLokasiKerjaList().forEach(l => { const n = String(l.namaLokasi || "").trim(); if (n && !seen[n]) { seen[n] = true; locations.push(n); } });
  } catch (e) {}
  return { locations: locations };
}


// MODUL LEMBUR (OVERTIME)
// -------------------------------------------------------------
function getLemburList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("PermintaanLembur");
  if (!sheet) return [];

  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];

  const result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({
      id: data[i][0],
      tglPengajuan: data[i][1],
      outlet: data[i][2],
      periode: data[i][3],
      idKaryawan: data[i][4],
      namaKaryawan: data[i][5],
      jabatan: data[i][6],
      managerApproval: data[i][7],
      hari: data[i][8],
      tglLembur: data[i][9],
      jamMasuk: data[i][10],
      jamKeluar: data[i][11],
      keterangan: data[i][12] || "-",
      statusApproval: data[i][13],
      catatanManager: data[i][14],
      createdBy: data[i][19] || ""
    });
  }
  return result;
}

function createLembur(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("PermintaanLembur");
  const lockC = LockService.getScriptLock();
  lockC.waitLock(20000);
  try {
  const id = nextLemburId_(sheet);
  const tglPengajuan = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  sheet.appendRow([
    id,
    tglPengajuan,
    data.outlet,
    data.periode,
    data.idKaryawan,
    data.namaKaryawan,
    data.jabatan,
    data.managerApproval,
    data.hari,
    data.tglLembur,
    data.jamMasuk,
    data.jamKeluar,
    data.keterangan,
    "Pending",
    "-",
    "",
    "",
    "0",
    "Rp 0",
    data.createdBy || ""
  ]);

  return { success: true, message: "Pengajuan lembur berhasil dibuat." };
  } finally { lockC.releaseLock(); }
}

/** ID berikutnya = nomor OVT terbesar + 1 (aman setelah ada baris yang dihapus). */
function nextLemburId_(sheet) {
  let max = 0;
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(function (r) {
      const m = /^OVT-(\d+)$/.exec(String(r[0]).trim());
      if (m) max = Math.max(max, Number(m[1]));
    });
  }
  return "OVT-" + String(max + 1).padStart(4, "0");
}

/**
 * Perbaiki ID lembur ganda (akibat ID lama dibuat dari jumlah baris sehingga bisa kembar setelah hapus data).
 * Baris pertama mempertahankan ID; baris kembar berikutnya diberi ID baru.
 * Dipanggil otomatis saat data lembur dimuat; aman dijalankan berulang.
 */
function perbaikiIdLemburDuplikat() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PermintaanLembur");
  if (!sheet || sheet.getLastRow() <= 1) return 0;
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); });
  const seen = {}; const dupRows = [];
  ids.forEach(function (id, i) { if (seen[id]) dupRows.push(i); else seen[id] = true; });
  if (!dupRows.length) return 0;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let max = 0;
    ids.forEach(function (id) { const m = /^OVT-(\d+)$/.exec(id); if (m) max = Math.max(max, Number(m[1])); });
    dupRows.forEach(function (i) {
      max++;
      sheet.getRange(i + 2, 1).setValue("OVT-" + String(max).padStart(4, "0"));
    });
  } finally { lock.releaseLock(); }
  return dupRows.length;
}


// -------------------------------------------------------------
// FUNGSI JEMBATAN (BRIDGE) UNTUK MODUL LEMBUR - PERBAIKAN KONEKSI
// index.html memanggil nama-nama fungsi di bawah ini; sebelumnya
// fungsi ini tidak ada di backend sehingga tombol Lembur tidak
// berfungsi (data gagal dimuat, RpNaN, Setujui/Tolak/Hapus gagal).
// -------------------------------------------------------------


// ---------------- REALISASI LEMBUR DARI ADMS (finger) ----------------
// Realisasi = (jam pulang - jam masuk finger ADMS) - 9 jam kerja; pembulatan per jam (sisa >= 31 menit naik).
const JAM_KERJA_NORMAL_MENIT_ = 9 * 60;

function loadAdmsHarianMap_() {
  const map = {};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AbsensiHarian");
  if (!sh || sh.getLastRow() <= 1) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach(function (r) {
    const id = String(r[0] || "").trim();
    const tgl = tglKeStr_(r[3]);
    if (!id || !tgl) return;
    map[id + "|" + tgl] = { masuk: jamKeStr_(r[4]).substring(0, 5), pulang: jamKeStr_(r[5]).replace(/\s*\(\+1\)/, "").substring(0, 5) };
  });
  return map;
}

function hitungLemburAdms_(masuk, pulang) {
  if (!/^\d{1,2}:\d{2}$/.test(masuk) || !/^\d{1,2}:\d{2}$/.test(pulang)) return null;
  const m1 = timeStringToMinutes_(masuk);
  let m2 = timeStringToMinutes_(pulang);
  if (m2 < m1) m2 += 24 * 60;
  const lebih = (m2 - m1) - JAM_KERJA_NORMAL_MENIT_;
  const jam = lebih > 0 ? Math.floor(lebih / 60) + ((lebih % 60) >= 31 ? 1 : 0) : 0;
  return { jam: jam, biaya: jam * 10000 };
}

/** Realisasi 1 pengajuan: murni dari finger ADMS (masuk & pulang pada tanggal lembur); tanpa data ADMS = belum ada realisasi. */
function resolveRealisasiLembur_(item, admsMap) {
  const a = admsMap[String(item.idKaryawan).trim() + "|" + String(item.tglLembur || "").substring(0, 10)];
  if (a && a.masuk && a.pulang) {
    const calc = hitungLemburAdms_(a.masuk, a.pulang);
    if (calc) return { jamText: a.masuk + " - " + a.pulang, jam: calc.jam, biaya: calc.biaya, sumber: "ADMS" };
  }
  return { jamText: "-", jam: 0, biaya: 0, sumber: "" };
}

/**
 * Menyediakan data Lembur dalam format yang dipakai renderLemburTable()
 * di index.html (field: outlet, manager, hariTanggal, jadwalJam, status,
 * realisasiJam, totalJam, totalBiaya sebagai angka).
 *
 * ROW-LEVEL SECURITY:
 * - Admin & HR: melihat seluruh data.
 * - Selain itu: hanya melihat data yang (a) manager approver-nya adalah
 *   dirinya sendiri, atau (b) dibuat oleh dirinya sendiri.
 */
function getLemburData(userRole, username) {
  try {
    try { perbaikiIdLemburDuplikat(); } catch (e) { /* abaikan */ }
    const role = (userRole || "").toString().trim();
    const uname = (username || "").toString().trim();
    const isFullAccess = (role === "Admin" || role === "HR");

    const lokasiOk = getAllowedLokasiForUser_(uname); // null = semua lokasi
    const list = getLemburList().filter(item => {
      if (lokasiOk) {
        const o = String(item.outlet || '').trim().toLowerCase();
        if (lokasiOk.names.indexOf(o) === -1 && lokasiOk.kodes.indexOf(o) === -1) return false;
      }
      if (isFullAccess) return true;
      if (item.managerApproval && item.managerApproval === uname) return true;
      if (item.createdBy && item.createdBy === uname) return true;
      return false;
    });

    const admsMap = loadAdmsHarianMap_();
    return list.map(item => {
      const jadwalJam = (item.jamMasuk || item.jamKeluar) ? `${item.jamMasuk || '-'} - ${item.jamKeluar || '-'}` : '-';
      const hariTanggal = (item.hari || item.tglLembur) ? `${item.hari || ''}${item.hari && item.tglLembur ? ', ' : ''}${item.tglLembur || ''}` : '-';

      const rr = resolveRealisasiLembur_(item, admsMap);
      const realisasiJam = rr.jamText;
      const totalJam = rr.jam;
      const totalBiaya = rr.biaya;

      return {
        id: item.id,
        tglPengajuan: item.tglPengajuan,
        idKaryawan: item.idKaryawan,
        namaKaryawan: item.namaKaryawan,
        jabatan: item.jabatan,
        outlet: item.outlet,
        periode: item.periode,
        manager: item.managerApproval,
        hariTanggal: hariTanggal,
        jadwalJam: jadwalJam,
        keterangan: item.keterangan,
        status: item.statusApproval,
        catatan: item.catatanManager && item.catatanManager !== '-' ? item.catatanManager : '',
        realisasiJam: realisasiJam,
        totalJam: totalJam,
        totalBiaya: totalBiaya,
        sumber: rr.sumber,
        createdBy: item.createdBy
      };
    });
  } catch (error) {
    Logger.log("Error getLemburData: " + error.message);
    return [];
  }
}

/**
 * Setujui / Tolak pengajuan lembur (dipanggil tombol Setujui/Tolak).
 * `catatan` dipakai terutama sebagai Alasan Penolakan saat status = "Ditolak",
 * tersimpan di kolom "Catatan Manager" (kolom O) pada sheet PermintaanLembur.
 */
function updateOvertimeStatus(id, newStatus, username, catatan) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("PermintaanLembur");
    if (!sheet) return { success: false, message: "Sheet PermintaanLembur tidak ditemukan." };

    const user = findUserByUsername_(username);
    const role = user ? String(user.role || "").trim() : "";
    const uname = String(username || "").trim();

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) !== String(id)) continue;
        const cur = String(data[i][13] || "Pending").trim().toUpperCase();
        const menungguHrd = cur.indexOf("HRD") > -1;
        const menungguManager = !menungguHrd && cur.indexOf("PENDING") > -1;
        if (!menungguHrd && !menungguManager) {
          return { success: false, message: "Pengajuan ini sudah berstatus final (" + data[i][13] + ")." };
        }
        const isAdmin = role === "Admin";
        const isHr = role === "HR";
        const isManagerTujuan = role === "Manager" && String(data[i][7] || "").trim().toLowerCase() === uname.toLowerCase();

        // Hak approval per tahap
        if (menungguManager && !(isManagerTujuan || isAdmin)) {
          return { success: false, message: "Tahap ini menunggu approval Manager yang dipilih (" + (data[i][7] || "-") + ")." };
        }
        if (menungguHrd && !(isHr || isAdmin)) {
          return { success: false, message: "Tahap ini menunggu approval HRD." };
        }

        let statusBaru, pesan;
        if (newStatus === "Ditolak") {
          statusBaru = "Ditolak";
          pesan = "Pengajuan lembur " + id + " ditolak.";
        } else if (menungguManager) {
          statusBaru = "Pending HRD";
          pesan = "Disetujui Manager. Pengajuan lembur " + id + " diteruskan ke HRD (status: Pending Approval HRD).";
        } else {
          statusBaru = "Disetujui";
          pesan = "Pengajuan lembur " + id + " disetujui HRD (final).";
        }
        sheet.getRange(i + 1, 14).setValue(statusBaru);
        // Jejak perubahan status: siapa & kapan (tahap Manager dipertahankan saat HRD memproses)
        const waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        const lama = String(data[i][14] || "").trim();
        const jejakLama = (menungguHrd && lama && lama !== "-") ? lama : "";
        const tahap = menungguHrd ? "HRD" : "Manager";
        let jejak;
        if (statusBaru === "Ditolak") {
          jejak = "Ditolak " + tahap + " oleh " + uname + " (" + waktu + ")" + (catatan && catatan !== "-" ? ": " + catatan : "");
        } else {
          jejak = "Disetujui " + tahap + " oleh " + uname + " (" + waktu + ")";
        }
        sheet.getRange(i + 1, 15).setValue(jejakLama ? jejakLama + " | " + jejak : jejak);
        return { success: true, message: pesan, status: statusBaru };
      }
      return { success: false, message: "Data lembur tidak ditemukan." };
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    return { success: false, message: "Gagal memperbarui status: " + error.message };
  }
}


/**
 * Hapus satu baris pengajuan lembur (tombol Hapus di Daftar Lembur & Realisasi).
 */
function deleteLemburData(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("PermintaanLembur");
    if (!sheet) return { success: false, message: "Sheet PermintaanLembur tidak ditemukan." };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Data lembur " + id + " berhasil dihapus." };
      }
    }
    return { success: false, message: "Data lembur tidak ditemukan." };
  } catch (error) {
    return { success: false, message: "Gagal menghapus data lembur: " + error.message };
  }
}




// -------------------------------------------------------------
// MODUL KARYAWAN
// -------------------------------------------------------------
function getKaryawanList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Karyawan");
  if (!sheet) return [];

  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];

  const result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({
      id: data[i][0],
      tanggalJoin: data[i][1],
      nama: data[i][2],
      email: data[i][3],
      noHp: data[i][4],
      penempatan: data[i][5],
      posisi: data[i][6],
      statusDokumen: data[i][7]
    });
  }
  return result;
}
























// ---------------- MASTER LOKASI KERJA (CRUD) ----------------

/**
 * Helper: ubah nilai sel jadi teks aman untuk dikirim ke browser. Objek Date/waktu mentah dari
 * getValues() sering gagal dikirim google.script.run (hasilnya null untuk seluruh array).
 */
function toClientText_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  return (v === null || v === undefined) ? "" : String(v);
}

function getLokasiKerjaList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MasterLokasiKerja");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      kodeLokasi: toClientText_(data[i][0]), namaLokasi: toClientText_(data[i][1]),
      jadwalKerja: toClientText_(data[i][2]), jamKerja: toClientText_(data[i][3]),
      hariOff: toClientText_(data[i][4]), pembagiHariKerja: Number(data[i][5]) || 25,
      tipeLokasi: toClientText_(data[i][6]) || "Store"
    });
  }
  return result;
}

/**
 * Simpan lokasi kerja baru, atau perbarui kalau Kode Lokasi sudah ada (upsert).
 */
function saveLokasiKerja(d) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("MasterLokasiKerja");
    if (!sheet) { setupPayrollModule(); sheet = ss.getSheetByName("MasterLokasiKerja"); }

    const kodeBaru = String(d.kodeLokasi || "").trim();
    const kodeLama = String(d.kodeLama || "").trim() || kodeBaru;
    if (!kodeBaru) return { success: false, message: "Kode Lokasi tidak boleh kosong." };

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const data = sheet.getDataRange().getValues();
      let rowLama = -1;
      for (let i = 1; i < data.length; i++) {
        const k = String(data[i][0]).trim();
        if (k === kodeLama) rowLama = i + 1;
        if (kodeBaru !== kodeLama && k.toLowerCase() === kodeBaru.toLowerCase()) {
          return { success: false, message: "Kode lokasi " + kodeBaru + " sudah dipakai lokasi lain." };
        }
      }

      if (rowLama > -1) {
        sheet.getRange(rowLama, 1, 1, 7).setNumberFormat("@").setValues([[
          kodeBaru, d.namaLokasi, d.jadwalKerja, d.jamKerja || "", d.hariOff || "", d.pembagiHariKerja || 25, d.tipeLokasi || "Store"
        ]]);
        if (kodeBaru !== kodeLama) {
          const n = gantiKodeLokasiDiSemuaData_(kodeLama, kodeBaru);
          return { success: true, message: "Kode lokasi diubah " + kodeLama + " -> " + kodeBaru + ". " + n.karyawan + " karyawan, " + n.absensi + " baris absensi, dan " + n.hakAkses + " pengaturan hak akses ikut diperbarui." };
        }
        return { success: true, message: "Lokasi kerja " + kodeBaru + " berhasil diperbarui." };
      }

      sheet.appendRow([kodeBaru, d.namaLokasi, d.jadwalKerja, d.jamKerja || "", d.hariOff || "", d.pembagiHariKerja || 25, d.tipeLokasi || "Store"]);
      return { success: true, message: "Lokasi kerja " + kodeBaru + " berhasil ditambahkan." };
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    return { success: false, message: "Gagal menyimpan lokasi kerja: " + error.message };
  }
}

/**
 * Import massal lokasi kerja dari template Excel. rows = [Kode Lokasi, Nama Lokasi, Jadwal Kerja, Jam Kerja, Hari Off, Pembagi Hari Kerja, Tipe Lokasi].
 * Kode yang sudah ada diperbarui, kode baru ditambahkan.
 */
function importLokasiBulk(rows, userRole) {
  const lock = LockService.getScriptLock();
  try {
    if (!isHrRole_(userRole)) return { success: false, message: "Akses ditolak. Import hanya untuk HR & Admin.", errors: [] };
    if (!rows || !rows.length) return { success: true, message: "Tidak ada lokasi di file.", errors: [], ins: 0, upd: 0 };
    lock.waitLock(20000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("MasterLokasiKerja");
    if (!sheet) { setupPayrollModule(); sheet = ss.getSheetByName("MasterLokasiKerja"); }

    const last = sheet.getLastRow();
    const existing = last > 1 ? sheet.getRange(2, 1, last - 1, 7).getValues() : [];
    const idx = {};
    existing.forEach((r, i) => { idx[String(r[0]).trim()] = i; });

    const errors = [], baru = [], seen = {};
    let ins = 0, upd = 0;
    rows.forEach((raw, n) => {
      const line = n + 2;
      const v = raw.map(x => String(x === null || x === undefined ? "" : x).trim());
      const kode = v[0], nama = v[1];
      if (!kode) { errors.push("Lokasi baris " + line + ": Kode Lokasi kosong."); return; }
      if (seen[kode.toLowerCase()]) { errors.push("Lokasi baris " + line + ": kode " + kode + " dobel di file - dilewati."); return; }
      seen[kode.toLowerCase()] = true;
      if (!nama) { errors.push("Lokasi baris " + line + " (" + kode + "): Nama Lokasi kosong."); return; }
      const tipe = /head|gudang|office/i.test(v[6]) ? "Head Office" : "Store";
      const pembagi = Number(v[5]) > 0 ? Number(v[5]) : 25;
      const rec = [kode, nama, v[2], v[3], v[4], pembagi, tipe];
      if (idx[kode] !== undefined) { existing[idx[kode]] = rec; upd++; } else { baru.push(rec); ins++; }
    });
    if (upd) sheet.getRange(2, 1, existing.length, 7).setNumberFormat("@").setValues(existing);
    if (baru.length) sheet.getRange(sheet.getLastRow() + 1, 1, baru.length, 7).setNumberFormat("@").setValues(baru);
    return { success: true, message: "Lokasi: " + ins + " ditambahkan, " + upd + " diperbarui.", errors: errors, ins: ins, upd: upd };
  } catch (err) {
    return { success: false, message: "Gagal import lokasi: " + err.message, errors: [] };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** Ganti Kode Lokasi lama -> baru di data karyawan, absensi harian, dan hak akses lokasi. Return jumlah baris terdampak. */
function gantiKodeLokasiDiSemuaData_(lama, baru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ganti = function (sheetName, col, syarat) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() <= 1) return 0;
    const n = sh.getLastRow() - 1;
    const vals = sh.getRange(2, col, n, 1).getValues();
    const cek = syarat ? sh.getRange(2, syarat.col, n, 1).getValues() : null;
    let hit = 0;
    for (let i = 0; i < n; i++) {
      if (String(vals[i][0]).trim() === lama && (!cek || String(cek[i][0]).trim() === syarat.nilai)) { vals[i][0] = baru; hit++; }
    }
    if (hit) sh.getRange(2, col, n, 1).setNumberFormat("@").setValues(vals);
    return hit;
  };
  return {
    karyawan: ganti("MasterGajiKaryawan", 3),
    absensi: ganti("AbsensiHarian", 3),
    hakAkses: ganti("Permissions", 5, { col: 4, nilai: "lokasi" })
  };
}

// ---------------- SKEMA KOMPONEN GAJI (Bulanan vs Harian) BERDASARKAN LEVEL JABATAN + TIPE LOKASI ----------------
// Aturan: Gaji Pokok selalu aktif. Tunjangan Jabatan/Makan Bulanan/Transport Bulanan HANYA untuk
// level Manager ke atas (semua lokasi) ATAU level SPV DI Head Office/Gudang Marlina.
// Tunjangan Makan Harian/Transport Harian untuk level Staff (semua lokasi) ATAU level SPV DI Store.
const JABATAN_LEVEL_MAP_ = {
  "Direktur": "Manager", "Logistik & Warehouse Manager": "Manager", "Area Manager": "Manager",
  "Chef": "SPV", "Trainer": "SPV",
  "Internal Audit Supervisor": "SPV", "Internal Audit Staff": "Staff",
  "Manager Finance & Accounting": "Manager", "Assistant Manager Finance & Accounting": "Manager",
  "Finance Supervisor": "SPV", "Accounting Supervisor": "SPV",
  "Account Payable Staff": "Staff", "Account Receivable Staff": "Staff",
  "Finance Accountig Tax Spv": "SPV", "Accounting Staff": "Staff", "Finance Staff": "Staff",
  "Kurir": "Staff", "Purchasing Manager": "Manager", "Purchasing Supervisor": "SPV",
  "Purchasing Senior Staff": "Staff", "Purchasing Staff": "Staff",
  "IT Support": "Staff", "Driver Pribadi": "Staff",
  "QA & QC Supervisor": "SPV", "QA & QC Staff": "Staff",
  "Project & Design Supervisor": "SPV", "Admin Project": "Staff",
  "Assistant Manager HRD": "Manager", "HRD Supervisor": "SPV", "HRD Staff": "Staff",
  "Office Girl": "Staff", "Assistant Manager Bussines Development": "Manager",
  "Marketing Manager": "Manager", "Marketing Supervisor": "SPV", "Marketing Staff": "Staff",
  "Design Grafis Staff": "Staff", "Marketing Senior Staff": "Staff",
  "Training Development Manager": "Manager", "Training Development Senior Staff": "Staff",
  "Training Development Staff": "Staff", "Personal Assistant": "Staff",
  "Admin Warehouse Senior Staff": "Staff", "Admin Logistik": "Staff", "Admin Checker": "Staff",
  "Admin Gudang": "Staff", "Gudang Staff": "Staff", "Driver": "Staff", "Teknisi Staff": "Staff",
  "Store Supervisor": "SPV", "Cashier Leader": "SPV", "Leader kitchen": "SPV",
  "Cook Helper Agemono": "Staff", "Cook Helper Sushi": "Staff", "Cashier": "Staff", "Stocker": "Staff"
};

/**
 * Jabatan dengan komponen HARIAN + Tunjangan Jabatan di lokasi mana pun (termasuk Head Office).
 * Cerminannya ada di index.html (JABATAN_HARIAN_DENGAN_TUNJ) - ubah di kedua tempat bila bertambah.
 */
const JABATAN_HARIAN_DENGAN_TUNJ_ = ["Trainer"];

/**
 * Skema komponen gaji:
 *  - Manager                      : Bulanan, ada Tunjangan Jabatan
 *  - SPV/Leader di Head Office    : Bulanan, ada Tunjangan Jabatan
 *  - SPV/Leader di Store          : Harian, ada Tunjangan Jabatan (Store Supervisor, Cashier Leader, Leader Kitchen, dst.)
 *  - Trainer (di lokasi mana pun) : Harian, ada Tunjangan Jabatan
 *  - Staff                        : Harian, tanpa Tunjangan Jabatan
 * Return {mode: "Bulanan"|"Harian", tunjJabatan: true|false}
 */
function getGajiSkema_(jabatan, tipeLokasi) {
  if (JABATAN_HARIAN_DENGAN_TUNJ_.indexOf(String(jabatan).trim()) > -1) return { mode: "Harian", tunjJabatan: true };
  const level = getLevelForJabatan_(jabatan);
  if (level === "Manager") return { mode: "Bulanan", tunjJabatan: true };
  if (level === "SPV") return { mode: tipeLokasi === "Head Office" ? "Bulanan" : "Harian", tunjJabatan: true };
  return { mode: "Harian", tunjJabatan: false };
}

function skemaLabel_(sk) { return sk.mode + (sk.mode === "Harian" && sk.tunjJabatan ? " + Tunjangan Jabatan" : ""); }

/**
 * Cari Level (Staff/SPV/Manager) untuk 1 Jabatan. Sumber kebenaran UTAMA adalah sheet
 * MasterLevelJabatan (bisa dikoreksi HR kapan saja tanpa ubah kode) - JABATAN_LEVEL_MAP_
 * cuma fallback kalau jabatan belum terdaftar di sheet (misalnya sheet baru pertama kali dibuat).
 */
function getLevelForJabatan_(jabatan) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MasterLevelJabatan");
  if (sheet && sheet.getLastRow() > 1) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(jabatan).trim()) {
        return data[i][1] || "Staff";
      }
    }
  }
  return JABATAN_LEVEL_MAP_[jabatan] || "Staff";
}

/**
 * Nolkan komponen gaji yang tidak berlaku sesuai skema Jabatan + Tipe Lokasi (jaga-jaga kalau
 * validasi di frontend ter-bypass). Gaji Pokok tidak pernah dinolkan.
 */
function enforceGajiScheme_(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shLokasi = ss.getSheetByName("MasterLokasiKerja");
  let tipeLokasi = "Store";
  if (shLokasi) {
    const lokasiData = shLokasi.getDataRange().getValues();
    for (let i = 1; i < lokasiData.length; i++) {
      if (String(lokasiData[i][0]).trim() === String(d.kodeLokasi).trim()) {
        tipeLokasi = lokasiData[i][6] || "Store";
        break;
      }
    }
  }
  const skema = getGajiSkema_(d.jabatan, tipeLokasi);
  if (skema.mode === "Bulanan") {
    d.tunjMakanHarian = 0;
    d.tunjTransportHarian = 0;
  } else {
    d.tunjMakanBulanan = 0;
    d.tunjTransportBulanan = 0;
    if (!skema.tunjJabatan) d.tunjJabatan = 0;
  }
  return d;
}

// ---------------- PAYROLL SETTINGS (Key-Value Global) ----------------






/** Daftar user (dengan password) untuk panel Manajemen User; hanya Admin/HR dengan sesi sah. */
function getUsersList(token, username) {
  requireHrSession_(token, username);
  return loadUsersInternal_();
}

function loadUsersInternal_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    if (!sheet) return [];

    const lockRepair = LockService.getScriptLock();
    if (lockRepair.tryLock(5000)) {
      try { repairDuplicateUserIds_(sheet); } finally { lockRepair.releaseLock(); }
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const users = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      
      users.push({
        id: row[0],
        username: row[1],
        email: row[2],
        password: row[3], // Disimpan untuk kebutuhan tampilan/edit Admin
        role: row[4],
        createdAt: row[5] ? Utilities.formatDate(new Date(row[5]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : "-"
      });
    }
    return users;
  } catch (error) {
    throw new Error("Gagal mengambil data user: " + error.message);
  }
}

  function saveNewUser(userData, token, requesterUsername) {
  const lock = LockService.getScriptLock();
  try {
    requireHrSession_(token, requesterUsername);
    lock.waitLock(20000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Users");
    
    if (!sheet) {
      setupUsersTable();
      sheet = ss.getSheetByName("Users");
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Validasi Duplikasi Username atau Email
    for (let i = 1; i < data.length; i++) {
      if (data[i][1].toString().toLowerCase() === userData.username.toLowerCase()) {
        return { success: false, message: "Username sudah digunakan!" };
      }
      if (data[i][2].toString().toLowerCase() === userData.email.toLowerCase()) {
        return { success: false, message: "Email sudah terdaftar!" };
      }
    }

    // Generate Auto ID
    // ID = nomor tertinggi yang ada + 1 (bukan jumlah baris, supaya tidak dobel setelah ada user yang dihapus)
    let maxNum = 0;
    for (let i = 1; i < data.length; i++) {
      const m = /^USR-(\d+)$/.exec(String(data[i][0]).trim());
      if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
    }
    const newId = "USR-" + String(maxNum + 1).padStart(3, '0');
    const createdAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    const newRow = [
      newId,
      userData.username.trim(),
      userData.email.trim(),
      userData.password,
      userData.role,
      createdAt
    ];
    
    sheet.appendRow(newRow);
    return { success: true, message: "User " + userData.username + " berhasil dibuat!" };
  } catch (error) {
    return { success: false, message: "Gagal menyimpan user: " + error.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** Perbaiki ID user yang dobel: baris pertama mempertahankan ID, baris berikutnya diberi ID baru (nomor tertinggi + 1). */
function repairDuplicateUserIds_(sheet) {
  const last = sheet.getLastRow();
  if (last < 3) return;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  const seen = {};
  let maxNum = 0, dup = false;
  ids.forEach(function (r) {
    const id = String(r[0]).trim();
    const m = /^USR-(\d+)$/.exec(id);
    if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
  });
  ids.forEach(function (r) {
    const id = String(r[0]).trim();
    if (!id) return;
    if (seen[id]) { maxNum++; r[0] = "USR-" + String(maxNum).padStart(3, "0"); dup = true; } else { seen[id] = true; }
  });
  if (dup) sheet.getRange(2, 1, ids.length, 1).setValues(ids);
}


  /**
 * Memperbarui Data User (Khusus Admin & HR)
 */
function updateUser(userData, token, requesterUsername) {
  const lock = LockService.getScriptLock();
  try {
    requireHrSession_(token, requesterUsername);
    lock.waitLock(20000);
    // Layanan Spreadsheet kadang gagal sesaat ("gagal saat mengakses dokumen"); coba ulang otomatis sampai 4 kali.
    return retrySpreadsheet_(function () {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName("Users");
      if (!sheet) return { success: false, message: "Sheet Users tidak ditemukan!" };

      const data = sheet.getDataRange().getValues();
      let rowIndex = -1;
      const uname = String(userData.username || "").trim().toLowerCase();
      const mail = String(userData.email || "").trim().toLowerCase();

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(userData.id)) {
          rowIndex = i + 1;
        } else {
          if (uname && String(data[i][1] || "").toLowerCase() === uname) return { success: false, message: "Username sudah digunakan oleh akun lain!" };
          if (mail && String(data[i][2] || "").toLowerCase() === mail) return { success: false, message: "Email sudah terdaftar pada akun lain!" };
        }
      }
      if (rowIndex === -1) return { success: false, message: "User ID tidak ditemukan!" };

      // Kolom B (Username), C (Email), D (Password), E (Role) ditulis dalam satu operasi
      sheet.getRange(rowIndex, 2, 1, 4).setValues([[String(userData.username).trim(), String(userData.email).trim(), userData.password, userData.role]]);
      SpreadsheetApp.flush();
      return { success: true, message: "Data user " + userData.username + " berhasil diperbarui!" };
    });
  } catch (error) {
    return { success: false, message: "Gagal memperbarui user: " + error.message + " (coba lagi beberapa saat)." };
  } finally {
    try { lock.releaseLock(); } catch (e) { /* abaikan */ }
  }
}

/** Jalankan fn; bila Layanan Spreadsheet gagal sementara, ulangi (maks 4x, jeda bertambah). */
function retrySpreadsheet_(fn) {
  let lastErr;
  for (let i = 0; i < 4; i++) {
    try { return fn(); } catch (e) {
      lastErr = e;
      if (String(e && e.message).indexOf("Spreadsheet") === -1 && String(e && e.message).indexOf("Layanan") === -1) throw e;
      Utilities.sleep(700 * (i + 1));
    }
  }
  throw lastErr;
}

/**
 * Menghapus User Berdasarkan ID (Khusus Admin & HR)
 */
/** Hapus semua aturan hak akses / lokasi per-User milik user yang dihapus, supaya tidak diwarisi akun baru ber-ID sama. */
function hapusAturanUser_(id, username) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Permissions");
  if (!sh || sh.getLastRow() < 2) return;
  const d = sh.getDataRange().getValues();
  for (let i = d.length - 1; i >= 1; i--) {
    if (d[i][1] === 'USER' && (String(d[i][2]).trim() === id || String(d[i][2]).trim().toLowerCase() === username.toLowerCase())) sh.deleteRow(i + 1);
  }
}

function deleteUser(userId, token, requesterUsername) {
  const lock = LockService.getScriptLock();
  try {
    const me = requireHrSession_(token, requesterUsername);
    lock.waitLock(20000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    if (!sheet) return { success: false, message: "Sheet Users tidak ditemukan!" };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(userId).trim()) {
        if (String(data[i][1]).trim().toLowerCase() === String(me.username).trim().toLowerCase()) {
          return { success: false, message: "Akun yang sedang dipakai tidak bisa dihapus." };
        }
        const idHapus = String(data[i][0]).trim(), unameHapus = String(data[i][1]).trim();
        sheet.deleteRow(i + 1);
        hapusAturanUser_(idHapus, unameHapus);
        return { success: true, message: "User berhasil dihapus dari sistem." };
      }
    }
    return { success: false, message: "User ID tidak ditemukan." };
  } catch (error) {
    return { success: false, message: "Gagal menghapus user: " + error.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

  /**
 * Mengambil daftar user dengan Role Manager/Admin untuk Dropdown Approver
 */
function getManagerList(callerRole) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const managers = [];
    
    for (let i = 1; i < data.length; i++) {
      const role = data[i][4] ? data[i][4].toString().trim() : "";
      const username = data[i][1] ? data[i][1].toString().trim() : "";
      const email = data[i][2] ? data[i][2].toString().trim() : "";
      
      const hanyaManager = String(callerRole || "") === "MOD"; // akun MOD hanya memilih level Manager
      if (role === "Manager" || (role === "Admin" && !hanyaManager)) {
        managers.push({
          username: username,
          email: email,
          label: username + " (" + role + ")"
        });
      }
    }
    return managers;
  } catch (error) {
    Logger.log("Error getManagerList: " + error.message);
    return [];
  }
}



// =============================================================
// MODUL PENGGAJIAN (PAYROLL)
// =============================================================

/**
 * Ambil 1 nilai konfigurasi global dari sheet PayrollSettings.
 */
function getPayrollSetting(key, defaultVal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("PayrollSettings");
  if (!sheet) return defaultVal;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(key).trim()) return data[i][1];
  }
  return defaultVal;
}









// ---------------- MASTER GAJI KARYAWAN (CRUD) ----------------

function getMasterGajiKaryawan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MasterGajiKaryawan");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();

  // Agregasi Cuti Tahunan yang sudah Approved (Status Final) di tahun berjalan, per NIK/ID Karyawan
  const tahunIni = new Date().getFullYear();
  const cutiMap = {};
  const shCuti = ss.getSheetByName("CutiIzin");
  if (shCuti && shCuti.getLastRow() > 1) {
    const cutiData = shCuti.getDataRange().getValues();
    for (let i = 1; i < cutiData.length; i++) {
      const nik = String(cutiData[i][2] || "").trim();
      const kategori = cutiData[i][7];
      const statusFinal = cutiData[i][24];
      if (!nik || kategori !== "Cuti Tahunan" || statusFinal !== "Approved") continue;
      const tgl = new Date(cutiData[i][1]);
      if (!isNaN(tgl.getTime()) && tgl.getFullYear() !== tahunIni) continue;
      cutiMap[nik] = (cutiMap[nik] || 0) + (Number(cutiData[i][10]) || 0);
    }
  }

  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const idKaryawan = String(data[i][0]).trim();
    const jatahCutiTahunan = data[i][17] !== "" && data[i][17] !== undefined ? Number(data[i][17]) : 12;
    const cutiTerpakai = cutiMap[idKaryawan] || 0;

    // PENTING: sel tanggal terbaca sebagai objek Date native oleh getValues(). google.script.run
    // sering gagal mengirim objek Date ke browser (client menerima null) - jadi harus dikonversi
    // ke teks biasa dulu sebelum dikembalikan ke frontend.
    let tglMasukVal = data[i][15];
    if (tglMasukVal instanceof Date) {
      tglMasukVal = Utilities.formatDate(tglMasukVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      tglMasukVal = tglMasukVal ? String(tglMasukVal) : "";
    }

    result.push({
      idKaryawan: data[i][0], nama: data[i][1], kodeLokasi: data[i][2], jabatan: data[i][3],
      statusKaryawan: data[i][4],
      gajiPokok: Number(data[i][5]) || 0, tunjJabatan: Number(data[i][6]) || 0,
      tunjMakanBulanan: Number(data[i][7]) || 0, tunjTransportBulanan: Number(data[i][8]) || 0,
      tunjMakanHarian: Number(data[i][9]) || 0, tunjTransportHarian: Number(data[i][10]) || 0,
      noRekening: data[i][11], namaBank: data[i][12], npwp: data[i][13], statusPtkp: data[i][14],
      tglMasuk: tglMasukVal, statusAktif: data[i][16] || "Aktif",
      jatahCutiTahunan: jatahCutiTahunan, cutiTerpakai: cutiTerpakai,
      sisaCuti: Math.max(0, jatahCutiTahunan - cutiTerpakai)
    });
  }
  return result;
}

function saveMasterGajiKaryawan(d) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("MasterGajiKaryawan");
    if (!sheet) { setupPayrollModule(); sheet = ss.getSheetByName("MasterGajiKaryawan"); }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(d.idKaryawan).trim()) {
        return { success: false, message: "ID Karyawan " + d.idKaryawan + " sudah terdaftar di Data Master Karyawan." };
      }
    }

    sheet.appendRow([
      d.idKaryawan, d.nama, d.kodeLokasi, d.jabatan, d.statusKaryawan || "Tetap",
      0, 0, 0, 0, 0, 0,
      "", "", "-", "-",
      "", d.statusAktif || "Aktif", 12
    ]);
    return { success: true, message: "Karyawan " + d.nama + " berhasil disimpan." };
  } catch (error) {
    return { success: false, message: "Gagal menyimpan karyawan: " + error.message };
  }
}

function updateMasterGajiKaryawan(d) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MasterGajiKaryawan");
    if (!sheet) return { success: false, message: "Sheet MasterGajiKaryawan tidak ditemukan." };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(d.idKaryawan).trim()) {
        const rowIndex = i + 1;
        const idBaru = String(d.idBaru || "").trim();
        if (idBaru && idBaru !== String(data[i][0]).trim()) {
          if (!/^EMP-/i.test(String(data[i][0]).trim())) return { success: false, message: "ID Karyawan hanya bisa diganti untuk ID sementara (EMP-...)." };
          for (let j = 1; j < data.length; j++) {
            if (String(data[j][0]).trim() === idBaru) return { success: false, message: "ID Karyawan " + idBaru + " sudah terdaftar di Data Master Karyawan." };
          }
          sheet.getRange(rowIndex, 1).setNumberFormat("@").setValue(idBaru);
        }
        // Hanya kolom identitas yang diubah; komponen gaji/bank/cuti yang sudah ada di sheet tidak disentuh
        sheet.getRange(rowIndex, 2, 1, 4).setValues([[d.nama, d.kodeLokasi, d.jabatan, d.statusKaryawan || data[i][4]]]);
        sheet.getRange(rowIndex, 17).setValue(d.statusAktif || "Aktif");
        return { success: true, message: "Karyawan " + d.nama + " berhasil diperbarui." };
      }
    }
    return { success: false, message: "ID Karyawan tidak ditemukan di Data Master Karyawan." };
  } catch (error) {
    return { success: false, message: "Gagal memperbarui karyawan: " + error.message };
  }
}

/**
 * Kembalikan daftar Jabatan -> Level (Staff/SPV/Manager) dari sheet MasterLevelJabatan,
 * dipakai frontend untuk auto-lock komponen gaji sesuai skema.
 */
function getLevelJabatanList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MasterLevelJabatan");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({ jabatan: data[i][0], level: data[i][1] || "Staff" });
  }
  return result;
}


// ---------------- PAYROLL INPUT (Absensi & Komponen Variabel) ----------------

// Pastikan header kolom tambahan (Potongan BPJS Manual, Event, Perjalanan Dinas, Pengganti Hari) ada di PayrollInput.
function ensurePayrollInputExtraHeaders_(sheet) {
  const heads = { 26: "Potongan BPJS Manual", 27: "Event (Hari)", 28: "Perjalanan Dinas / Luar Kota (Hari)", 29: "Pengganti Hari / PH (Hari)", 30: "Cuti Khusus (Hari)" };
  Object.keys(heads).forEach(function (col) {
    const cell = sheet.getRange(1, Number(col));
    if (!cell.getValue()) cell.setValue(heads[col]).setFontWeight("bold").setBackground("#008bca").setFontColor("#ffffff");
  });
}

function getPayrollInput(bulan, tahun) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("PayrollInput");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][1]) !== Number(bulan) || Number(data[i][2]) !== Number(tahun)) continue;

    // PENTING: kolom "Tanggal Input" (timestamp) otomatis dikonversi Google Sheets jadi objek Date.
    // google.script.run gagal mengirim objek Date mentah ke browser (hasilnya null untuk SELURUH
    // array), jadi harus dijadikan teks dulu sebelum dikembalikan.
    let tglInputVal = data[i][24];
    if (tglInputVal instanceof Date) {
      tglInputVal = Utilities.formatDate(tglInputVal, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    } else {
      tglInputVal = tglInputVal ? String(tglInputVal) : "";
    }

    result.push({
      idInput: data[i][0], bulan: Number(data[i][1]) || 0, tahun: Number(data[i][2]) || 0, idKaryawan: data[i][3],
      hariKerjaAktual: Number(data[i][4]) || 0, jumlahAlpa: Number(data[i][5]) || 0,
      jumlahSakitTanpaSurat: Number(data[i][6]) || 0, jumlahIzin: Number(data[i][7]) || 0,
      jumlahCuti: Number(data[i][8]) || 0, jumlahSakitDenganSurat: Number(data[i][9]) || 0,
      terlambat: Number(data[i][10]) || 0, pulangCepat: Number(data[i][11]) || 0,
      uangKerajinan: Number(data[i][12]) || 0, tunjToko: Number(data[i][13]) || 0,
      tunjMultiskill: Number(data[i][14]) || 0, pendapatanLainnya: Number(data[i][15]) || 0,
      rapelManual: Number(data[i][16]) || 0, potonganVarian: Number(data[i][17]) || 0,
      depositSeragam: Number(data[i][18]) || 0, depositTraining: Number(data[i][19]) || 0,
      potonganKaryawan: Number(data[i][20]) || 0, potonganPph21: Number(data[i][21]) || 0,
      catatan: data[i][22] || "", inputBy: data[i][23] || "", tglInput: tglInputVal,
      potonganBpjsManual: Number(data[i][25]) || 0,
      event: Number(data[i][26]) || 0, perjalananDinas: Number(data[i][27]) || 0, penggantiHari: Number(data[i][28]) || 0, cutiKhusus: Number(data[i][29]) || 0
    });
  }
  return result;
}

/**
 * Simpan / update input absensi & komponen variabel 1 karyawan untuk 1 periode.
 * Kalau data untuk ID Karyawan + Bulan + Tahun yang sama sudah ada, akan di-update (bukan dobel).
 */
function savePayrollInput(d) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("PayrollInput");
    if (!sheet) { setupPayrollModule(); sheet = ss.getSheetByName("PayrollInput"); }

    ensurePayrollInputExtraHeaders_(sheet);
    const data = sheet.getDataRange().getValues();
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    const extraValues = [d.event || 0, d.perjalananDinas || 0, d.penggantiHari || 0, d.cutiKhusus || 0];

    const rowValues = [
      d.hariKerjaAktual || 0, d.jumlahAlpa || 0, d.jumlahSakitTanpaSurat || 0, d.jumlahIzin || 0,
      d.jumlahCuti || 0, d.jumlahSakitDenganSurat || 0, d.terlambat || 0, d.pulangCepat || 0,
      d.uangKerajinan || 0, d.tunjToko || 0, d.tunjMultiskill || 0, d.pendapatanLainnya || 0, d.rapelManual || 0,
      d.potonganVarian || 0, d.depositSeragam || 0, d.depositTraining || 0, d.potonganKaryawan || 0,
      d.potonganPph21 || 0, d.catatan || "", d.inputBy || "", now, d.potonganBpjsManual || 0
    ];

    for (let i = 1; i < data.length; i++) {
      if (Number(data[i][1]) === Number(d.bulan) && Number(data[i][2]) === Number(d.tahun) &&
          String(data[i][3]).trim() === String(d.idKaryawan).trim()) {
        sheet.getRange(i + 1, 5, 1, 22).setValues([rowValues]);
        sheet.getRange(i + 1, 27, 1, 4).setValues([extraValues]);
        return { success: true, message: "Data absensi & komponen variabel berhasil diperbarui." };
      }
    }

    const idInput = `PIN-${d.tahun}${String(d.bulan).padStart(2, '0')}-${d.idKaryawan}`;
    sheet.appendRow([idInput, d.bulan, d.tahun, d.idKaryawan, ...rowValues, ...extraValues]);
    return { success: true, message: "Data absensi & komponen variabel berhasil disimpan." };
  } catch (error) {
    return { success: false, message: "Gagal menyimpan input payroll: " + error.message };
  }
}


// =============================================================
// INPUT DATA BULANAN PER HARI (checklist harian per karyawan)
// Disimpan di sheet "PayrollInputHarian"; total 1 periode (cutoff tgl 16 - 15) otomatis
// dikalkulasikan ke sheet PayrollInput (dipakai engine payroll) setiap kali disimpan.
// =============================================================
const INPUT_HARIAN_STATUS_ = ["Hadir", "Alpa", "Sakit Tanpa Surat", "Izin", "Cuti", "Cuti Khusus", "Sakit Dengan Surat"];

function getInputHarianSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("PayrollInputHarian");
  if (!sh) {
    sh = ss.insertSheet("PayrollInputHarian");
    sh.appendRow(["ID Karyawan", "Tanggal", "Bulan", "Tahun", "Status", "Terlambat", "Pulang Cepat", "Event",
                  "Perjalanan Dinas", "Pengganti Hari (PH)", "Input By", "Waktu Input"]);
    sh.getRange(1, 1, 1, 12).setFontWeight("bold").setBackground("#008bca").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.getRange("A:B").setNumberFormat("@");
  }
  return sh;
}

function isTrue_(v) { return v === true || String(v).toUpperCase() === "TRUE"; }

function readInputHarianRows_(sh) {
  const out = [];
  if (!sh || sh.getLastRow() <= 1) return out;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues();
  data.forEach(function (r, i) {
    if (!r[0]) return;
    out.push({
      rowNum: i + 2, id: String(r[0]).trim(), tanggal: tglKeStr_(r[1]),
      bulan: Number(r[2]) || 0, tahun: Number(r[3]) || 0, status: String(r[4] || "").trim(),
      terlambat: isTrue_(r[5]), pulangCepat: isTrue_(r[6]), event: isTrue_(r[7]), dinas: isTrue_(r[8]), ph: isTrue_(r[9])
    });
  });
  return out;
}

function hitungTotalInputHarian_(hari) {
  const t = { hadir: 0, alpa: 0, sakitTanpaSurat: 0, izin: 0, cuti: 0, cutiKhusus: 0, sakitDenganSurat: 0, terlambat: 0, pulangCepat: 0, event: 0, perjalananDinas: 0, penggantiHari: 0 };
  hari.forEach(function (d) {
    if (d.status === "Hadir") t.hadir++;
    else if (d.status === "Alpa") t.alpa++;
    else if (d.status === "Sakit Tanpa Surat") t.sakitTanpaSurat++;
    else if (d.status === "Izin") t.izin++;
    else if (d.status === "Cuti") t.cuti++;
    else if (d.status === "Cuti Khusus") t.cutiKhusus++;
    else if (d.status === "Sakit Dengan Surat") t.sakitDenganSurat++;
    if (d.terlambat) t.terlambat++;
    if (d.pulangCepat) t.pulangCepat++;
    if (d.event) t.event++;
    if (d.dinas) t.perjalananDinas++;
    if (d.ph) t.penggantiHari++;
  });
  return t;
}

/** Data harian 1 karyawan untuk 1 periode + data ADMS per tanggal + total bulanan lama (jika ada). */
function getInputHarianKaryawan(idKaryawan, bulan, tahun) {
  const id = String(idKaryawan).trim();
  const range = getPayrollPeriodeRange_(bulan, tahun);
  const hari = readInputHarianRows_(getInputHarianSheet_())
    .filter(function (r) { return r.id === id && r.bulan === Number(bulan) && r.tahun === Number(tahun); })
    .map(function (r) { return { tanggal: r.tanggal, status: r.status, terlambat: r.terlambat, pulangCepat: r.pulangCepat, event: r.event, dinas: r.dinas, ph: r.ph }; });
  const bulanan = getPayrollInput(bulan, tahun).filter(function (x) { return String(x.idKaryawan).trim() === id; })[0] || null;
  return { mulai: range.mulai, selesai: range.selesai, hari: hari, adms: getAbsensiDetail(id, bulan, tahun), bulanan: bulanan };
}

/** Inti penyimpanan: ganti seluruh baris harian karyawan di periode itu, lalu hitung ulang total ke PayrollInput. */
function persistInputHarian_(idKaryawan, bulan, tahun, hariList, inputBy, catatan) {
  const id = String(idKaryawan).trim();
  const range = getPayrollPeriodeRange_(bulan, tahun);
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  const bersih = [];
  const seen = {};
  (hariList || []).forEach(function (d) {
    const tgl = String(d.tanggal || "").substring(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(tgl) || tgl < range.mulai || tgl > range.selesai || seen[tgl]) return;
    const status = INPUT_HARIAN_STATUS_.indexOf(d.status) > -1 ? d.status : "";
    const item = { tanggal: tgl, status: status, terlambat: !!d.terlambat, pulangCepat: !!d.pulangCepat, event: !!d.event, dinas: !!d.dinas, ph: !!d.ph };
    if (!item.status && !item.terlambat && !item.pulangCepat && !item.event && !item.dinas && !item.ph) return;
    seen[tgl] = true;
    bersih.push(item);
  });
  bersih.sort(function (a, b) { return a.tanggal.localeCompare(b.tanggal); });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getInputHarianSheet_();
    const lama = readInputHarianRows_(sh).filter(function (r) { return r.id === id && r.bulan === Number(bulan) && r.tahun === Number(tahun); });
    for (let i = lama.length - 1; i >= 0; i--) sh.deleteRow(lama[i].rowNum);

    if (bersih.length) {
      const rows = bersih.map(function (d) {
        return [id, d.tanggal, Number(bulan), Number(tahun), d.status, d.terlambat, d.pulangCepat, d.event, d.dinas, d.ph, inputBy || "", now];
      });
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
    }

    const t = hitungTotalInputHarian_(bersih);
    const existing = getPayrollInput(bulan, tahun).filter(function (x) { return String(x.idKaryawan).trim() === id; })[0] || {};
    const merged = {};
    Object.keys(existing).forEach(function (k) { merged[k] = existing[k]; });
    merged.bulan = Number(bulan); merged.tahun = Number(tahun); merged.idKaryawan = id;
    merged.hariKerjaAktual = t.hadir; merged.jumlahAlpa = t.alpa; merged.jumlahSakitTanpaSurat = t.sakitTanpaSurat;
    merged.jumlahIzin = t.izin; merged.jumlahCuti = t.cuti; merged.cutiKhusus = t.cutiKhusus; merged.jumlahSakitDenganSurat = t.sakitDenganSurat;
    merged.terlambat = t.terlambat; merged.pulangCepat = t.pulangCepat;
    merged.event = t.event; merged.perjalananDinas = t.perjalananDinas; merged.penggantiHari = t.penggantiHari;
    merged.catatan = (catatan !== undefined && catatan !== null) ? catatan : (existing.catatan || "");
    merged.inputBy = inputBy || "";
    const res = savePayrollInput(merged);
    if (!res.success) throw new Error(res.message);
    return { success: true, totals: t, jumlahHari: bersih.length };
  } finally {
    lock.releaseLock();
  }
}

function saveInputHarian(d) {
  try {
    // Hadir tanpa data ADMS hanya boleh dicentang Admin/HR
    const pengguna = findUserByUsername_(d.inputBy);
    const peran = pengguna ? String(pengguna.role || "").trim() : "";
    if (peran !== "Admin" && peran !== "HR") {
      const admsTgl = {};
      getAbsensiDetail(d.idKaryawan, d.bulan, d.tahun).forEach(function (a) { admsTgl[a.tanggal] = true; });
      const sudahHadir = {};
      readInputHarianRows_(getInputHarianSheet_()).forEach(function (r) {
        if (r.id === String(d.idKaryawan).trim() && r.bulan === Number(d.bulan) && r.tahun === Number(d.tahun) && r.status === "Hadir") sudahHadir[r.tanggal] = true;
      });
      const langgar = (d.hari || []).filter(function (x) { return x.status === "Hadir" && !admsTgl[x.tanggal] && !sudahHadir[x.tanggal]; }).map(function (x) { return x.tanggal; });
      if (langgar.length) return { success: false, message: "Hadir tanpa data ADMS hanya bisa dicentang oleh Admin/HRD. Tanggal: " + langgar.join(", ") };
    }
    const r = persistInputHarian_(d.idKaryawan, d.bulan, d.tahun, d.hari, d.inputBy, d.catatan);
    return { success: true, message: "Input harian tersimpan (" + r.jumlahHari + " hari terisi). Total bulan: Hadir " + r.totals.hadir +
      ", Alpa " + r.totals.alpa + ", Izin " + r.totals.izin + ", Cuti " + r.totals.cuti + ", Cuti Khusus " + r.totals.cutiKhusus + ".", totals: r.totals };
  } catch (error) {
    return { success: false, message: "Gagal menyimpan input harian: " + error.message };
  }
}

/** Set ID karyawan yang sudah punya input harian di periode ini. */
function getIdsWithInputHarian_(bulan, tahun) {
  const set = {};
  readInputHarianRows_(getInputHarianSheet_()).forEach(function (r) {
    if (r.bulan === Number(bulan) && r.tahun === Number(tahun)) set[r.id] = true;
  });
  return set;
}

/** Gabungkan data ADMS ke input harian karyawan: hanya mengisi hari yang statusnya masih kosong. */
function mergeAdmsIntoInputHarian_(idKaryawan, bulan, tahun, inputBy) {
  const id = String(idKaryawan).trim();
  const adms = getAbsensiDetail(id, bulan, tahun);
  const byTgl = {};
  readInputHarianRows_(getInputHarianSheet_())
    .filter(function (r) { return r.id === id && r.bulan === Number(bulan) && r.tahun === Number(tahun); })
    .forEach(function (r) { byTgl[r.tanggal] = { tanggal: r.tanggal, status: r.status, terlambat: r.terlambat, pulangCepat: r.pulangCepat, event: r.event, dinas: r.dinas, ph: r.ph }; });
  adms.forEach(function (a) {
    const cur = byTgl[a.tanggal] || { tanggal: a.tanggal, status: "", terlambat: false, pulangCepat: false, event: false, dinas: false, ph: false };
    if (!cur.status) { cur.status = "Hadir"; if (a.status === "Terlambat") cur.terlambat = true; }
    byTgl[a.tanggal] = cur;
  });
  const list = Object.keys(byTgl).map(function (k) { return byTgl[k]; });
  return persistInputHarian_(id, bulan, tahun, list, inputBy, undefined);
}

// ---------------- ENGINE KALKULASI GAJI ----------------








// =============================================================
// MODUL ABSENSI - INTEGRASI ADMS/iCLOCK (MySQL via JDBC)
// =============================================================

/**
 * Baca konfigurasi bridge ADMS dari Script Properties (Project Settings > Script Properties).
 * Property yang dibutuhkan: ADMS_BRIDGE_URL (contoh: http://103.56.148.91:8899),
 * ADMS_BRIDGE_KEY (API key yang sama dengan yang diisi di absensi_bridge.py).
 */
function getAdmsBridgeConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    url: props.getProperty("ADMS_BRIDGE_URL"),
    key: props.getProperty("ADMS_BRIDGE_KEY")
  };
}

/**
 * FUNGSI TES - jalankan manual dari Editor untuk memastikan bridge Absensi di VPS bisa
 * dijangkau dari Apps Script sebelum dipakai sync beneran.
 */
function testKoneksiADMS() {
  const cfg = getAdmsBridgeConfig_();
  if (!cfg.url || !cfg.key) {
    Logger.log("GAGAL: Script Properties ADMS_BRIDGE_URL/ADMS_BRIDGE_KEY belum diisi.");
    return;
  }
  try {
    const resp = UrlFetchApp.fetch(cfg.url + "/ping?key=" + encodeURIComponent(cfg.key), {
      muteHttpExceptions: true, method: "get"
    });
    Logger.log("HTTP Status: " + resp.getResponseCode());
    Logger.log("Isi respons: " + resp.getContentText());
  } catch (err) {
    Logger.log("GAGAL connect ke bridge ADMS: " + err.message);
  }
}

/**
 * Sync absensi dari ADMS lewat bridge (endpoint /scan = scan mentah per karyawan).
 * Scan dikelompokkan menjadi SESI kerja, bukan per tanggal kalender, sehingga shift malam yang
 * melewati tengah malam tetap satu baris (tanggal = tanggal masuk, jam pulang diberi tanda "(+1)").
 *  - Scan berjarak < 5 menit dianggap tap ganda dan diabaikan.
 *  - Sesi baru dimulai bila jeda antar scan > PayrollSettings "AbsensiJedaSesiJam" (default 14 jam).
 *  - Sesi yang melebihi "AbsensiMaxDurasiJam" (default 16 jam, mis. pulang malam lalu masuk pagi)
 *    dianggap ambigu dan dipecah per tanggal kalender.
 *  - Sesi dengan 1 scan = "Scan Tunggal": tetap dihitung hadir (bukan 0 jam kerja), jam masuk/pulang
 *    dikosongkan sesuai jenis scan-nya, dan tidak dianggap terlambat kecuali jelas scan masuk HO.
 * Baris lama di rentang tanggal yang sama dihapus dulu (idempotent).
 * tglMulaiStr/tglSelesaiStr format "yyyy-MM-dd".
 */
/** Ubah daftar kode lokasi (array) jadi set huruf besar; null/undefined = tanpa filter (semua lokasi). */
function kodeFilterSet_(list) {
  if (list === null || list === undefined) return null;
  const set = {};
  list.forEach(function (k) { set[String(k).trim().toUpperCase()] = true; });
  return set;
}

/** kodeFilter (opsional) = array kode lokasi yang disinkronkan; kosong/null = semua lokasi. */
function syncAbsensiFromADMS(tglMulaiStr, tglSelesaiStr, kodeFilter) {
  try {
    const filterSet = kodeFilterSet_(kodeFilter);
    const lokasiOk = function (kode) { return !filterSet || !!filterSet[String(kode).trim().toUpperCase()]; };
    const cfg = getAdmsBridgeConfig_();
    if (!cfg.url || !cfg.key) {
      return { success: false, message: "Script Properties ADMS_BRIDGE_URL/ADMS_BRIDGE_KEY belum diisi.", jumlahBaris: 0, errors: [] };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let shAbsensi = ss.getSheetByName("AbsensiHarian");
    if (!shAbsensi) { setupAbsensiModule(); shAbsensi = ss.getSheetByName("AbsensiHarian"); }

    const masterMap = {};
    getMasterGajiKaryawan().forEach(k => { masterMap[normBadge_(k.idKaryawan)] = k; });
    const lokasiMap = {};
    getLokasiKerjaList().forEach(l => { lokasiMap[l.kodeLokasi] = l; });

    // Tarik 1 hari lebih lebar di kedua sisi supaya shift malam di tepi periode tetap terpasang
    const tz = Session.getScriptTimeZone();
    const geser = (s, n) => {
      const p = s.split("-");
      return Utilities.formatDate(new Date(+p[0], +p[1] - 1, +p[2] + n), tz, "yyyy-MM-dd");
    };
    const fetchUrl = cfg.url + "/scan?mulai=" + encodeURIComponent(geser(tglMulaiStr, -1)) +
      "&selesai=" + encodeURIComponent(geser(tglSelesaiStr, 1)) + "&key=" + encodeURIComponent(cfg.key);
    const resp = UrlFetchApp.fetch(fetchUrl, { muteHttpExceptions: true, method: "get" });

    if (resp.getResponseCode() === 404) {
      return { success: false, message: "Bridge ADMS belum diperbarui (endpoint /scan tidak ada). Terapkan patch absensi_bridge.py lalu jalankan ulang bridge.", jumlahBaris: 0, errors: [] };
    }
    if (resp.getResponseCode() !== 200) {
      return { success: false, message: "Bridge ADMS merespons error (" + resp.getResponseCode() + "): " + resp.getContentText(), jumlahBaris: 0, errors: [] };
    }

    const scans = JSON.parse(resp.getContentText()).data || [];
    const sesiList = buildSesiAbsensi_(scans);

    const rows = [];
    const errors = [];
    const badgeTercatat = {};
    let jumlahTunggal = 0, adaMesin = false;
    const now = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

    sesiList.forEach(s => {
      if (s.tanggal < tglMulaiStr || s.tanggal > tglSelesaiStr) return;
      const master = masterMap[normBadge_(s.employeeId)];
      if (master && !lokasiOk(master.kodeLokasi)) return; // di luar lokasi yang dipilih
      if (!master) {
        if (filterSet) return; // badge tak dikenal tidak bisa dipastikan lokasinya; dilaporkan hanya pada sync semua lokasi
        if (!badgeTercatat[s.employeeId]) {
          badgeTercatat[s.employeeId] = true;
          errors.push("Badge '" + s.employeeId + "' (" + s.nama + ") tidak ditemukan di Data Master Karyawan - dilewati.");
        }
        return;
      }

      const lokasi = lokasiMap[master.kodeLokasi];
      const tipeLokasi = lokasi ? lokasi.tipeLokasi : "Store";
      const jk = parseJamKerja_(lokasi && lokasi.jamKerja);
      let jamMasuk = s.masuk, jamPulang = s.pulang, status = "Hadir", terlambatMenit = 0;

      if (s.tunggal) {
        jumlahTunggal++;
        // Tentukan scan tunggal ini scan masuk atau pulang: sebelum/sesudah titik tengah jam kerja (default 12:00)
        const tengah = jk ? (timeStringToMinutes_(jk.mulai) + timeStringToMinutes_(jk.selesai)) / 2 : 720;
        if (timeStringToMinutes_(s.masuk) < tengah) { jamPulang = ""; } else { jamPulang = s.masuk; jamMasuk = ""; }
        status = "Scan Tunggal";
      }

      // Telat hanya dihitung untuk Head Office (Store belum punya jadwal shift per orang per hari)
      if (tipeLokasi === "Head Office" && jamMasuk) {
        const standar = timeStringToMinutes_(jk ? jk.mulai : "09:00");
        const aktual = timeStringToMinutes_(jamMasuk);
        if (aktual > standar) { status = "Terlambat"; terlambatMenit = aktual - standar; }
      }

      // Mesin asal scan (kalau bridge mengirim info mesin): sesi tunggal hanya punya satu mesin
      let mMasuk = s.mesinMasuk, mPulang = s.mesinPulang;
      if (s.tunggal) { if (jamMasuk) { mPulang = ""; } else { mMasuk = ""; mPulang = s.mesinMasuk; } }
      if (mMasuk || mPulang) adaMesin = true;
      rows.push([master.idKaryawan, master.nama, master.kodeLokasi, s.tanggal, jamMasuk, jamPulang, status, terlambatMenit, now, mMasuk || "", mPulang || ""]);
    });

    // Hapus baris lama di rentang tanggal yang sama (dan, bila ada filter, hanya untuk lokasi terpilih), per blok berurutan dari bawah
    const lastRow = shAbsensi.getLastRow();
    if (lastRow > 1) {
      const kolom = shAbsensi.getRange(2, 3, lastRow - 1, 2).getValues(); // C = Kode Lokasi, D = Tanggal
      let runEnd = -1;
      for (let i = kolom.length - 1; i >= -1; i--) {
        let hapus = false;
        if (i >= 0) {
          const t = tglKeStr_(kolom[i][1]);
          hapus = t >= tglMulaiStr && t <= tglSelesaiStr && lokasiOk(kolom[i][0]);
        }
        if (hapus && runEnd < 0) runEnd = i;
        if (!hapus && runEnd >= 0) { shAbsensi.deleteRows(i + 3, runEnd - i); runEnd = -1; }
      }
    }

    if (rows.length > 0) {
      const mulaiBaris = shAbsensi.getLastRow() + 1;
      // ID, tanggal, dan jam harus tetap teks (nol di depan badge jangan hilang, tanggal tidak jadi Date)
      [1, 4, 5, 6].forEach(c => shAbsensi.getRange(mulaiBaris, c, rows.length, 1).setNumberFormat("@"));
      shAbsensi.getRange(mulaiBaris, 1, rows.length, 11).setValues(rows);
      if (adaMesin) shAbsensi.getRange(1, 10, 1, 2).setValues([["Mesin Masuk", "Mesin Pulang"]]).setFontWeight("bold");
    }

    return {
      success: true,
      message: "Sync selesai: " + rows.length + " baris absensi (" + tglMulaiStr + " s.d. " + tglSelesaiStr + ")" +
        (jumlahTunggal ? ", " + jumlahTunggal + " di antaranya scan tunggal." : "."),
      jumlahBaris: rows.length,
      errors: errors
    };
  } catch (err) {
    return { success: false, message: "Gagal sync ADMS: " + err.message, jumlahBaris: 0, errors: [] };
  }
}

/**
 * Samakan format ID untuk pencocokan badge ADMS dengan ID Karyawan di Master Gaji.
 * ADMS menyimpan badge 9 digit berawalan nol (mis. "000202333"), sedangkan Master Gaji sering
 * berisi "202333" (nol di depan hilang, terutama kalau sel terbaca sebagai angka). Dibandingkan
 * tanpa nol di depan dan tanpa spasi/huruf besar-kecil.
 */
function normBadge_(v) {
  const s = String(v === null || v === undefined ? "" : v).trim().toUpperCase();
  return /^\d+$/.test(s) ? (s.replace(/^0+/, "") || "0") : s;
}

/** Kelompokkan scan mentah ({employee_id, nama, waktu}) menjadi sesi kerja per karyawan. */
function buildSesiAbsensi_(scans) {
  const jedaMs = (Number(getPayrollSetting("AbsensiJedaSesiJam", 14)) || 14) * 3600000;
  const maxMs = (Number(getPayrollSetting("AbsensiMaxDurasiJam", 16)) || 16) * 3600000;
  const dupMs = 5 * 60000;
  const toMs = w => {
    const m = String(w).match(/(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})/);
    return m ? new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]).getTime() : NaN;
  };

  // Scan dari SEMUA mesin digabung per karyawan: badge dinormalisasi (000202333 = 202333) supaya
  // masuk di mesin A dan pulang di mesin B tetap menjadi 1 sesi kerja.
  const byEmp = {};
  scans.forEach(s => {
    const rawId = String(s.employee_id || "").trim();
    const id = normBadge_(rawId);
    const t = toMs(s.waktu);
    if (!id || isNaN(t)) return;
    const mesin = String(s.sn || s.device_sn || s.mesin || s.device || s.terminal || s.alias || "").trim();
    const e = (byEmp[id] = byEmp[id] || { nama: s.nama, rawId: rawId, list: [] });
    e.list.push({ t: t, m: mesin, str: String(s.waktu).replace("T", " ").substring(0, 19) });
  });

  const result = [];
  Object.keys(byEmp).forEach(id => {
    const nama = byEmp[id].nama;
    const mk = grup => {
      const f = grup[0], l = grup[grup.length - 1];
      const tunggal = grup.length === 1;
      let pulang = tunggal ? "" : l.str.substring(11, 19);
      if (!tunggal && l.str.substring(0, 10) !== f.str.substring(0, 10)) pulang += " (+1)";
      result.push({ employeeId: byEmp[id].rawId, nama: nama, tanggal: f.str.substring(0, 10), masuk: f.str.substring(11, 19), pulang: pulang, tunggal: tunggal,
                    mesinMasuk: f.m || "", mesinPulang: tunggal ? "" : (l.m || "") });
    };
    const tutup = grup => {
      if (!grup) return;
      if (grup.length > 1 && grup[grup.length - 1].t - grup[0].t > maxMs) {
        const per = {};
        grup.forEach(s => { const k = s.str.substring(0, 10); (per[k] = per[k] || []).push(s); });
        Object.keys(per).sort().forEach(k => mk(per[k]));
      } else {
        mk(grup);
      }
    };

    let cur = null, lastT = 0;
    byEmp[id].list.sort((a, b) => a.t - b.t).forEach(s => {
      if (cur && s.t - lastT < dupMs) return; // tap ganda
      if (!cur || s.t - lastT > jedaMs) { tutup(cur); cur = [s]; } else { cur.push(s); }
      lastT = s.t;
    });
    tutup(cur);
  });
  return result;
}

/** "09:00-18:00" -> {mulai:"09:00", selesai:"18:00"}; null kalau formatnya tidak terbaca. */
function parseJamKerja_(jamKerja) {
  const m = String(jamKerja || "").match(/(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})/);
  return m ? { mulai: m[1], selesai: m[2] } : null;
}

/** Sel tanggal bisa terbaca sebagai Date oleh getValues(); samakan ke "yyyy-MM-dd". */
function tglKeStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(v === null || v === undefined ? "" : v).substring(0, 10);
}

/** Sel jam bisa terbaca sebagai Date; samakan ke "HH:mm:ss" (teks biasa dikembalikan apa adanya). */
function jamKeStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "HH:mm:ss");
  return String(v === null || v === undefined ? "" : v);
}

function timeStringToMinutes_(hhmm) {
  const parts = String(hhmm).split(":");
  return (parseInt(parts[0], 10) || 0) * 60 + (parseInt(parts[1], 10) || 0);
}

/**
 * Rekap absensi per karyawan untuk 1 periode payroll (cutoff tanggal 16 bulan sebelumnya
 * s.d. tanggal 15 bulan berjalan), dikelompokkan per Kode Lokasi di frontend.
 */
function getAbsensiRekap(bulan, tahun) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("AbsensiHarian");
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const range = getPayrollPeriodeRange_(bulan, tahun);
  const data = sheet.getDataRange().getValues();

  const rekapMap = {};
  for (let i = 1; i < data.length; i++) {
    const tglStr = tglKeStr_(data[i][3]);
    if (tglStr < range.mulai || tglStr > range.selesai) continue;

    const idKaryawan = String(data[i][0]).trim();
    if (!rekapMap[idKaryawan]) {
      rekapMap[idKaryawan] = {
        idKaryawan: idKaryawan, nama: data[i][1], kodeLokasi: data[i][2],
        hariHadir: 0, hariTerlambat: 0, totalMenitTerlambat: 0, hariScanTunggal: 0
      };
    }
    rekapMap[idKaryawan].hariHadir++;
    if (data[i][6] === "Scan Tunggal") rekapMap[idKaryawan].hariScanTunggal++;
    if (data[i][6] === "Terlambat") {
      rekapMap[idKaryawan].hariTerlambat++;
      rekapMap[idKaryawan].totalMenitTerlambat += Number(data[i][7]) || 0;
    }
  }

  return Object.keys(rekapMap).map(k => rekapMap[k]);
}

/**
 * Detail absensi harian 1 karyawan untuk 1 periode payroll (untuk drill-down di UI).
 */
function getAbsensiDetail(idKaryawan, bulan, tahun) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("AbsensiHarian");
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const range = getPayrollPeriodeRange_(bulan, tahun);
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== String(idKaryawan).trim()) continue;
    const tglStr = tglKeStr_(data[i][3]);
    if (tglStr < range.mulai || tglStr > range.selesai) continue;
    result.push({
      tanggal: tglStr, jamMasuk: jamKeStr_(data[i][4]), jamPulang: jamKeStr_(data[i][5]),
      status: data[i][6], terlambatMenit: Number(data[i][7]) || 0
    });
  }
  return result.sort((a, b) => a.tanggal.localeCompare(b.tanggal));
}

/**
 * Hitung rentang cutoff payroll (tgl 16 bulan sebelumnya s.d. tgl 15 bulan berjalan) sebagai
 * string "yyyy-MM-dd", dipakai untuk filter AbsensiHarian.
 */
function getPayrollPeriodeRange_(bulan, tahun) {
  const akhir = new Date(tahun, bulan - 1, 15);
  const mulai = new Date(tahun, bulan - 2, 16);
  const fmt = d => Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return { mulai: fmt(mulai), selesai: fmt(akhir) };
}

/**
 * Tarik rekap absensi (Hari Kerja Aktual + Datang Terlambat) ke sheet PayrollInput untuk
 * 1 periode. HANYA 2 kolom ini yang ditimpa - field lain yang sudah diisi HR manual (Alpa,
 * Izin, Cuti, Kerajinan, PPh21, dst) TIDAK disentuh. Kalau baris PayrollInput untuk karyawan
 * itu belum ada, baris baru dibuat dengan field lain default 0.
 */
function tarikAbsensiKePayrollInput(bulan, tahun, inputBy, kodeFilter) {
  try {
    const filterSet = kodeFilterSet_(kodeFilter);
    const rekap = getAbsensiRekap(bulan, tahun).filter(function (r) { return !filterSet || !!filterSet[String(r.kodeLokasi).trim().toUpperCase()]; });
    if (!rekap.length) {
      return { success: false, message: "Belum ada data absensi tersinkron untuk periode ini. Sync dari ADMS dahulu.", jumlahDiupdate: 0 };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("PayrollInput");
    if (!sheet) { setupPayrollModule(); sheet = ss.getSheetByName("PayrollInput"); }

    const data = sheet.getDataRange().getValues();
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    let jumlahDiupdate = 0;
    const dailyIds = getIdsWithInputHarian_(bulan, tahun); // karyawan yang sudah diinput harian

    rekap.forEach(r => {
      if (dailyIds[String(r.idKaryawan).trim()]) {
        mergeAdmsIntoInputHarian_(r.idKaryawan, bulan, tahun, inputBy);
        jumlahDiupdate++;
        return;
      }
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (Number(data[i][1]) === Number(bulan) && Number(data[i][2]) === Number(tahun) &&
            String(data[i][3]).trim() === String(r.idKaryawan).trim()) {
          const rowNum = i + 1;
          sheet.getRange(rowNum, 5).setValue(r.hariHadir);               // Hari Kerja Aktual
          sheet.getRange(rowNum, 11).setValue(r.hariTerlambat);          // Datang Terlambat (Kali)
          sheet.getRange(rowNum, 25).setValue(now);                     // Tanggal Input
          found = true;
          jumlahDiupdate++;
          break;
        }
      }
      if (!found) {
        const idInput = "PIN-" + tahun + String(bulan).padStart(2, "0") + "-" + r.idKaryawan;
        sheet.appendRow([
          idInput, bulan, tahun, r.idKaryawan,
          r.hariHadir, 0, 0, 0, 0, 0, r.hariTerlambat, 0,
          0, 0, 0, 0, 0,
          0, 0, 0, 0, 0,
          "Ditarik otomatis dari ADMS", inputBy || "", now, 0
        ]);
        jumlahDiupdate++;
      }
    });

    return { success: true, message: jumlahDiupdate + " karyawan berhasil diperbarui di Input Data Bulanan.", jumlahDiupdate: jumlahDiupdate };
  } catch (err) {
    return { success: false, message: "Gagal menarik data: " + err.message, jumlahDiupdate: 0 };
  }
}


// =============================================================
// MUTASI LOKASI + IMPORT/EXPORT MASTER GAJI KARYAWAN
// =============================================================

/** Guard hapus: sesi sah, Admin/HR, fitur master-karyawan.<fitur> tidak dimatikan, dan lokasi target boleh diakses akun. */
function guardHapusMaster_(token, username, fitur, kodeLokasi) {
  const me = requireHrSession_(token, username);
  const perms = computeEffectivePermissions_(me.id, me.username, me.role);
  if (perms['master-karyawan.' + fitur] === false) throw new Error("Akses ditolak. Akun ini tidak memiliki izin " + (fitur === 'hapus_lokasi' ? "hapus lokasi." : "hapus karyawan."));
  const allowed = computeEffectiveLokasi_(me.id, me.username, me.role);
  if (allowed !== null && kodeLokasi !== undefined) {
    const set = kodeFilterSet_(allowed);
    if (!set[String(kodeLokasi).trim().toUpperCase()]) throw new Error("Lokasi " + kodeLokasi + " di luar hak akses lokasi akun ini.");
  }
  return me;
}

/** Hapus 1 karyawan dari Data Master Karyawan. Riwayat absensi/lembur/input bulanan yang sudah tercatat tidak ikut dihapus. */
function deleteMasterKaryawan(idKaryawan, token, username) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MasterGajiKaryawan");
    if (!sheet) return { success: false, message: "Sheet MasterGajiKaryawan tidak ditemukan." };
    const data = sheet.getDataRange().getValues();
    const id = String(idKaryawan).trim();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === id) {
        guardHapusMaster_(token, username, 'hapus_karyawan', String(data[i][2]).trim());
        const nama = data[i][1];
        sheet.deleteRow(i + 1);
        return { success: true, message: "Karyawan " + nama + " (" + id + ") dihapus dari Data Master Karyawan." };
      }
    }
    guardHapusMaster_(token, username, 'hapus_karyawan');
    return { success: false, message: "Karyawan dengan ID " + id + " tidak ditemukan." };
  } catch (err) {
    return { success: false, message: "Gagal menghapus karyawan: " + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** Hapus 1 lokasi kerja. Hanya boleh bila sudah tidak ada karyawan di lokasi itu. */
function deleteMasterLokasi(kodeLokasi, token, username) {
  const lock = LockService.getScriptLock();
  try {
    const kode = String(kodeLokasi).trim();
    guardHapusMaster_(token, username, 'hapus_lokasi', kode);
    lock.waitLock(20000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shLok = ss.getSheetByName("MasterLokasiKerja");
    if (!shLok) return { success: false, message: "Sheet MasterLokasiKerja tidak ditemukan." };

    const shKar = ss.getSheetByName("MasterGajiKaryawan");
    if (shKar && shKar.getLastRow() > 1) {
      const kodeKar = shKar.getRange(2, 3, shKar.getLastRow() - 1, 1).getValues();
      const jumlah = kodeKar.filter(function (r) { return String(r[0]).trim() === kode; }).length;
      if (jumlah > 0) {
        return { success: false, message: "Lokasi " + kode + " masih punya " + jumlah + " karyawan. Pindahkan (Mutasi) atau hapus karyawannya dulu." };
      }
    }

    const data = shLok.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === kode) {
        const nama = data[i][1];
        shLok.deleteRow(i + 1);
        // Bersihkan aturan akses lokasi yang menunjuk ke kode ini
        const shPerm = ss.getSheetByName("Permissions");
        if (shPerm && shPerm.getLastRow() > 1) {
          const p = shPerm.getDataRange().getValues();
          for (let j = p.length - 1; j >= 1; j--) {
            if (p[j][3] === 'lokasi' && String(p[j][4]).trim() === kode) shPerm.deleteRow(j + 1);
          }
        }
        return { success: true, message: "Lokasi " + kode + " - " + nama + " dihapus." };
      }
    }
    return { success: false, message: "Lokasi " + kode + " tidak ditemukan." };
  } catch (err) {
    return { success: false, message: "Gagal menghapus lokasi: " + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function isHrRole_(role) {
  return ["HR", "Admin", "HR Manager", "HR Leader"].indexOf(role) > -1;
}

/** Sheet riwayat mutasi (dibuat otomatis saat pertama dipakai). */
function getMutasiSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("MutasiLokasi");
  if (!sh) {
    sh = ss.insertSheet("MutasiLokasi");
    sh.appendRow(["ID Mutasi", "Tanggal Efektif", "ID Karyawan", "Nama", "Lokasi Asal", "Lokasi Tujuan",
                  "Jabatan Lama", "Jabatan Baru", "Alasan", "Diinput Oleh", "Waktu Input"]);
    sh.getRange(1, 1, 1, 11).setFontWeight("bold").setBackground("#008bca").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    ["B:B", "C:C", "K:K"].forEach(r => sh.getRange(r).setNumberFormat("@"));
  }
  return sh;
}

function appendMutasiRows_(rows) {
  if (!rows.length) return;
  const sh = getMutasiSheet_();
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 11).setValues(rows);
}

/**
 * Mutasi 1 karyawan ke lokasi lain. Langsung berlaku (tanpa approval) dan dicatat di MutasiLokasi.
 * d: {idKaryawan, kodeLokasiBaru, jabatanBaru (opsional), tanggalEfektif (opsional), alasan}
 * Komponen gaji yang tidak berlaku di skema baru (Bulanan/Harian) dinolkan sesuai aturan aplikasi.
 */
function mutasiKaryawan(d, userRole, username) {
  const lock = LockService.getScriptLock();
  try {
    if (!isHrRole_(userRole)) return { success: false, message: "Akses ditolak. Mutasi hanya untuk HR & Admin." };
    if (!d || !d.idKaryawan || !d.kodeLokasiBaru) return { success: false, message: "Karyawan dan lokasi tujuan wajib diisi." };
    lock.waitLock(20000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MasterGajiKaryawan");
    if (!sheet) return { success: false, message: "Sheet MasterGajiKaryawan tidak ditemukan." };

    const lokasiMap = {};
    getLokasiKerjaList().forEach(l => { lokasiMap[String(l.kodeLokasi).trim()] = l; });
    const kodeBaru = String(d.kodeLokasiBaru).trim();
    if (!lokasiMap[kodeBaru]) return { success: false, message: "Kode Lokasi tujuan '" + kodeBaru + "' tidak ada di Master Lokasi Kerja." };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== String(d.idKaryawan).trim()) continue;

      const kodeLama = String(data[i][2]).trim();
      const jabLama = String(data[i][3]).trim();
      const jabBaru = String(d.jabatanBaru || jabLama).trim();
      if (kodeLama === kodeBaru) return { success: false, message: "Lokasi tujuan sama dengan lokasi saat ini." };

      const tipeLama = lokasiMap[kodeLama] ? lokasiMap[kodeLama].tipeLokasi : "Store";
      const skemaLama = skemaLabel_(getGajiSkema_(jabLama, tipeLama));
      const skemaBaru = skemaLabel_(getGajiSkema_(jabBaru, lokasiMap[kodeBaru].tipeLokasi));

      const komp = enforceGajiScheme_({
        kodeLokasi: kodeBaru, jabatan: jabBaru,
        tunjJabatan: Number(data[i][6]) || 0, tunjMakanBulanan: Number(data[i][7]) || 0,
        tunjTransportBulanan: Number(data[i][8]) || 0, tunjMakanHarian: Number(data[i][9]) || 0,
        tunjTransportHarian: Number(data[i][10]) || 0
      });
      sheet.getRange(i + 1, 3, 1, 2).setValues([[kodeBaru, jabBaru]]);
      sheet.getRange(i + 1, 7, 1, 5).setValues([[komp.tunjJabatan, komp.tunjMakanBulanan, komp.tunjTransportBulanan, komp.tunjMakanHarian, komp.tunjTransportHarian]]);

      const tz = Session.getScriptTimeZone();
      const now = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
      const efektif = d.tanggalEfektif || Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
      appendMutasiRows_([["MUT-" + Date.now(), efektif, String(data[i][0]).trim(), data[i][1], kodeLama, kodeBaru,
                          jabLama, jabBaru, d.alasan || "", username || "", now]]);

      let msg = data[i][1] + " dimutasi dari " + kodeLama + " ke " + kodeBaru + ". Perubahan langsung berlaku.";
      if (skemaLama !== skemaBaru) {
        msg += "\n\nPERHATIAN: skema gaji berubah dari " + skemaLama + " ke " + skemaBaru +
          ". Komponen tunjangan skema lama dinolkan; cek dan isi komponen yang berlaku lewat tombol Edit.";
      }
      return { success: true, message: msg };
    }
    return { success: false, message: "ID Karyawan tidak ditemukan di Data Master Karyawan." };
  } catch (err) {
    return { success: false, message: "Gagal mutasi: " + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** Riwayat mutasi (terbaru di atas). idKaryawan opsional untuk memfilter 1 karyawan. */
function getMutasiList(idKaryawan) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MutasiLokasi");
  if (!sh || sh.getLastRow() <= 1) return [];
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues();
  const out = [];
  for (let i = data.length - 1; i >= 0; i--) {
    if (idKaryawan && String(data[i][2]).trim() !== String(idKaryawan).trim()) continue;
    out.push({
      id: toClientText_(data[i][0]), tanggalEfektif: toClientText_(data[i][1]), idKaryawan: toClientText_(data[i][2]),
      nama: toClientText_(data[i][3]), lokasiAsal: toClientText_(data[i][4]), lokasiTujuan: toClientText_(data[i][5]),
      jabatanLama: toClientText_(data[i][6]), jabatanBaru: toClientText_(data[i][7]), alasan: toClientText_(data[i][8]),
      diinputOleh: toClientText_(data[i][9]), waktuInput: toClientText_(data[i][10])
    });
    if (out.length >= 300) break;
  }
  return out;
}

/**
 * Import massal Data Master Karyawan dari CSV. rows = array baris (string) dengan urutan kolom tetap:
 * [ID Karyawan, Nama, Kode Lokasi, Jabatan, Status Karyawan, Status Aktif].
 * ID yang sudah ada diperbarui (hanya kolom itu; gaji/bank/cuti yang ada di sheet tidak disentuh), ID baru ditambahkan.
 */
function importMasterGajiBulk(rows, userRole, username) {
  const lock = LockService.getScriptLock();
  try {
    if (!isHrRole_(userRole)) return { success: false, message: "Akses ditolak. Import hanya untuk HR & Admin.", errors: [] };
    if (!rows || !rows.length) return { success: false, message: "Tidak ada baris data untuk diimport.", errors: [] };
    lock.waitLock(20000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("MasterGajiKaryawan");
    if (!sheet) { setupPayrollModule(); sheet = ss.getSheetByName("MasterGajiKaryawan"); }
    sheet.getRange("A:A").setNumberFormat("@"); // ID harus teks supaya nol di depan tidak hilang

    const tz = Session.getScriptTimeZone();
    const now = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
    const hariIni = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    const lokasiMap = {};
    getLokasiKerjaList().forEach(l => { lokasiMap[String(l.kodeLokasi).trim()] = l; });

    const lastRow = sheet.getLastRow();
    const existing = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 18).getValues() : [];
    const idx = {};
    existing.forEach((r, i) => { idx[String(r[0]).trim()] = i; });

    const errors = [], newRows = [], mutasi = [], seen = {};
    let ins = 0, upd = 0;

    rows.forEach((raw, n) => {
      const line = n + 2; // nomor baris di file (baris 1 = header)
      const c = raw.map(v => String(v === null || v === undefined ? "" : v).trim());
      let id = c[0];
      if (!id) { errors.push("Baris " + line + ": ID Karyawan kosong."); return; }
      // ID numerik yang kehilangan nol di depan (efek Excel) dicocokkan ke ID 9 digit yang sudah ada
      if (idx[id] === undefined && /^\d+$/.test(id) && id.length < 9 && idx[("000000000" + id).slice(-9)] !== undefined) id = ("000000000" + id).slice(-9);
      if (seen[id]) { errors.push("Baris " + line + ": ID " + id + " dobel di file - dilewati."); return; }
      seen[id] = true;

      const nama = c[1], kode = c[2], jab = c[3];
      const statusKaryawan = c[4] || "Tetap", statusAktif = c[5] || "Aktif";
      if (!nama) { errors.push("Baris " + line + " (" + id + "): Nama kosong."); return; }
      if (!lokasiMap[kode]) { errors.push("Baris " + line + " (" + id + "): Kode Lokasi '" + kode + "' tidak ada di Master Lokasi Kerja."); return; }
      if (!jab) { errors.push("Baris " + line + " (" + id + "): Jabatan kosong."); return; }

      if (idx[id] !== undefined) {
        const lama = existing[idx[id]];
        if (String(lama[2]).trim() !== kode) {
          mutasi.push(["MUT-" + Date.now() + "-" + mutasi.length, hariIni, id, nama, String(lama[2]).trim(), kode,
                       String(lama[3]).trim(), jab, "Via import CSV", username || "", now]);
        }
        lama[1] = nama; lama[2] = kode; lama[3] = jab; lama[4] = statusKaryawan; lama[16] = statusAktif;
        upd++;
      } else {
        newRows.push([id, nama, kode, jab, statusKaryawan, 0, 0, 0, 0, 0, 0, "", "", "-", "-", "", statusAktif, 12]);
        ins++;
      }
    });

    if (upd > 0) sheet.getRange(2, 1, existing.length, 18).setValues(existing);
    if (newRows.length) sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 18).setValues(newRows);
    appendMutasiRows_(mutasi);

    return {
      success: (ins + upd) > 0,
      message: "Import selesai: " + ins + " karyawan ditambahkan, " + upd + " diperbarui, " + mutasi.length +
        " mutasi lokasi tercatat, " + errors.length + " baris dilewati.",
      errors: errors
    };
  } catch (err) {
    return { success: false, message: "Gagal import: " + err.message, errors: [] };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// =============================================================
// KARYAWAN BARU (HIRED) -> OTOMATIS MASUK MASTER GAJI SESUAI LOKASI PENEMPATAN
// =============================================================

/**
 * Buat baris draft di MasterGajiKaryawan: ID sementara EMP-xxxx, lokasi sesuai penempatan,
 * semua nominal 0, Status Aktif "Menunggu Data" (belum ikut payroll sampai HR melengkapi).
 * Return true bila baris dibuat, false bila sudah ada / dilewati.
 */
function buatDraftMasterGaji_(empId, nama, penempatan, posisi, tglJoin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("MasterGajiKaryawan");
  if (!sheet) { setupPayrollModule(); sheet = ss.getSheetByName("MasterGajiKaryawan"); }
  const ada = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(v => String(v).trim()) : [];
  if (ada.indexOf(String(empId).trim()) > -1) return false;

  const nm = String(penempatan || "").trim().toLowerCase();
  const lok = getLokasiKerjaList().filter(l => String(l.namaLokasi).trim().toLowerCase() === nm)[0];
  sheet.appendRow([String(empId).trim(), nama, lok ? lok.kodeLokasi : "", posisi || "", "Probation",
                   0, 0, 0, 0, 0, 0, "", "", "-", "-", tglJoin || "", "Menunggu Data", 12]);
  return true;
}

/** Tarik karyawan di sheet Karyawan yang belum ada di Master Gaji (cek ID EMP dan nama). */
function sinkronKaryawanBaru(userRole) {
  try {
    if (!isHrRole_(userRole)) return { success: false, message: "Akses ditolak. Hanya untuk HR & Admin." };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shEmp = ss.getSheetByName("Karyawan");
    if (!shEmp || shEmp.getLastRow() <= 1) return { success: true, message: "Belum ada data di Data Karyawan." };
    const namaAda = {};
    getMasterGajiKaryawan().forEach(k => { namaAda[String(k.nama).trim().toLowerCase()] = true; });
    let dibuat = 0, dilewati = 0;
    shEmp.getDataRange().getDisplayValues().slice(1).forEach(r => {
      const nama = String(r[2]).trim();
      if (!r[0] || !nama || namaAda[nama.toLowerCase()]) { dilewati++; return; }
      if (buatDraftMasterGaji_(r[0], nama, r[5], r[6], r[1])) { dibuat++; namaAda[nama.toLowerCase()] = true; } else dilewati++;
    });
    return { success: true, message: dibuat + " karyawan baru masuk ke Data Master Karyawan (status Menunggu Data), " + dilewati + " dilewati karena sudah ada." };
  } catch (err) {
    return { success: false, message: "Gagal sinkron: " + err.message };
  }
}

/** Sync absensi dari mesin ADMS untuk periode payroll, lalu isi Hari Kerja Aktual & Terlambat di PayrollInput. */
function syncDanTarikAdms(bulan, tahun, inputBy, kodeLokasiList) {
  try {
    // Lokasi yang diproses = pilihan user, dibatasi hak akses lokasi akunnya (null = semua lokasi)
    const me = findUserByUsername_(inputBy);
    const allowed = me ? computeEffectiveLokasi_(me.id, me.username, me.role) : [];
    let kode = (kodeLokasiList === undefined || kodeLokasiList === null) ? null : kodeLokasiList;
    if (allowed !== null) {
      const aset = kodeFilterSet_(allowed);
      kode = (kode === null ? allowed : kode.filter(function (k) { return aset[String(k).trim().toUpperCase()]; }));
    }
    if (kode !== null && kode.length === 0) return { success: false, message: "Tidak ada lokasi yang boleh disinkronkan untuk akun ini.", errors: [] };

    const range = getPayrollPeriodeRange_(bulan, tahun);
    const s = syncAbsensiFromADMS(range.mulai, range.selesai, kode);
    if (!s.success) return s;
    const t = tarikAbsensiKePayrollInput(bulan, tahun, inputBy, kode);
    const info = kode === null ? "Semua lokasi" : (kode.length === 1 ? "Lokasi " + kode[0] : kode.length + " lokasi");
    return { success: t.success, message: "[" + info + "] " + s.message + "\n" + t.message, errors: s.errors || [], jumlahDiupdate: t.jumlahDiupdate };
  } catch (err) {
    return { success: false, message: "Gagal sync & tarik dari ADMS: " + err.message, errors: [] };
  }
}


