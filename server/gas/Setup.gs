/**
 * Setup.gs - Inisialisasi Database Spreadsheet (Users, Karyawan, PermintaanLembur, Master Lokasi/Gaji, Absensi, Permissions)
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Tab Users
  let shUsers = ss.getSheetByName("Users");
  if (!shUsers) {
    shUsers = ss.insertSheet("Users");
    shUsers.appendRow(["Username", "Email", "Password", "Role"]);
    shUsers.appendRow(["admin", "admin@company.com", "admin123", "Admin"]);
    shUsers.appendRow(["hr_manager", "hr@company.com", "hr123", "HR"]);
    shUsers.appendRow(["Dede", "dede@company.com", "hr123", "HR"]);
    shUsers.appendRow(["Sherly", "sherly@company.com", "hr123", "HR"]);
    shUsers.appendRow(["Adinda", "adinda@company.com", "hr123", "HR"]);
    shUsers.appendRow(["Ika", "ika@company.com", "hr123", "HR"]);
    shUsers.appendRow(["mgr_ops", "manager.ops@company.com", "mgr123", "Manager"]);
    shUsers.appendRow(["mod_store", "mod.store@company.com", "mod123", "MOD"]);
  }

  // 2. Tab Karyawan
  let shEmp = ss.getSheetByName("Karyawan");
  if (!shEmp) {
    shEmp = ss.insertSheet("Karyawan");
    shEmp.appendRow(["ID Karyawan", "Tanggal Join", "Nama", "Email", "No. HP", "Penempatan", "Posisi", "Status Dokumen"]);
  }

  // 3. Tab PermintaanLembur
  let shOvt = ss.getSheetByName("PermintaanLembur");
  if (!shOvt) {
    shOvt = ss.insertSheet("PermintaanLembur");
    shOvt.appendRow([
      "ID Lembur", "Tanggal Pengajuan", "Outlet/Penempatan", "Periode Lembur",
      "ID Karyawan", "Nama Karyawan", "Jabatan/Posisi", "Manager Approval",
      "Hari", "Tanggal Lembur", "Jam Masuk (Jadwal)", "Jam Keluar (Jadwal)",
      "Keterangan", "Status Approval", "Catatan Manager", "Realisasi Mulai", "Realisasi Selesai",
      "Total Jam Lembur", "Total Biaya Lembur"
    ]);
  }

  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet) ss.deleteSheet(defaultSheet);
}


  /**
 * Inisialisasi & Auto-Repair Sheet Users (Standar 6 Kolom)
 */
function setupUsersTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Users");
  const headers = ["ID", "Username", "Email", "Password", "Role", "CreatedAt"];
  
  // 1. Jika Sheet "Users" BELUM ADA, buat sheet baru
  if (!sheet) {
    sheet = ss.insertSheet("Users");
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight("bold")
         .setBackground("#008bca")
         .setFontColor("#ffffff");
    
    // Tambahkan User Default Admin
    const defaultAdmin = [
      "USR-001",
      "admin",
      "admin@midelimakmurjaya.co.id",
      "admin123",
      "Admin",
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
    ];
    sheet.appendRow(defaultAdmin);
    sheet.setFrozenRows(1);
    Logger.log("Sheet Users berhasil dibuat dari awal.");
    return;
  }

  // 2. Jika Sheet SUDAH ADA, periksa apakah kolom A adalah "ID"
  const firstCell = sheet.getRange("A1").getValue().toString().trim();
  
  // Jika kolom A bukan ID (misal masih Username), sisipkan kolom ID di Kolom A secara otomatis
  if (firstCell.toLowerCase() !== "id") {
    sheet.insertColumnBefore(1);
    
    // Auto-generate ID untuk semua baris data lama yang belum memiliki ID
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      for (let i = 2; i <= lastRow; i++) {
        const generatedId = "USR-" + String(i - 1).padStart(3, '0');
        sheet.getRange(i, 1).setValue(generatedId);
      }
    }
  }

  // 3. Ketatkan/Standarkan Header & Tampilan (Baris 1)
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
       .setFontWeight("bold")
       .setBackground("#008bca")
       .setFontColor("#ffffff");
  sheet.setFrozenRows(1);
  
  Logger.log("Sheet Users berhasil disesuaikan secara otomatis ke standar 6 kolom.");
}

/**
 * =============================================================
 * MODUL PENGGAJIAN (PAYROLL) - Inisialisasi Sheet
 * =============================================================
 * Jalankan fungsi ini sekali dari Script Editor (pilih fungsi ini
 * di dropdown, lalu klik Run) untuk membuat seluruh sheet yang
 * dibutuhkan modul penggajian.
 */
/**
 * Auto-repair: tambahkan sheet MasterLevelJabatan (mapping Jabatan -> Level untuk
 * aturan komponen gaji bulanan/harian), dan tambahkan+isi kolom "Tipe Lokasi" di
 * MasterLokasiKerja kalau belum ada (Head Office / Store). Aman dijalankan berkali-kali.
 */
function setupLevelJabatanDanTipeLokasi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. MasterLevelJabatan
  let shLevel = ss.getSheetByName("MasterLevelJabatan");
  if (!shLevel) {
    shLevel = ss.insertSheet("MasterLevelJabatan");
    shLevel.appendRow(["Jabatan", "Level"]);
    shLevel.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#008bca").setFontColor("#ffffff");
    shLevel.setFrozenRows(1);

    const levelSeed = [
      ["Direktur", "Manager"], ["Logistik & Warehouse Manager", "Manager"], ["Area Manager", "Manager"],
      ["Chef", "SPV"], ["Trainer", "SPV"], ["Internal Audit Supervisor", "SPV"], ["Internal Audit Staff", "Staff"],
      ["Manager Finance & Accounting", "Manager"], ["Assistant Manager Finance & Accounting", "Manager"],
      ["Finance Supervisor", "SPV"], ["Accounting Supervisor", "SPV"], ["Account Payable Staff", "Staff"],
      ["Account Receivable Staff", "Staff"], ["Finance Accountig Tax Spv", "SPV"], ["Accounting Staff", "Staff"],
      ["Finance Staff", "Staff"], ["Kurir", "Staff"], ["Purchasing Manager", "Manager"],
      ["Purchasing Supervisor", "SPV"], ["Purchasing Senior Staff", "Staff"], ["Purchasing Staff", "Staff"],
      ["IT Support", "Staff"], ["Driver Pribadi", "Staff"], ["QA & QC Supervisor", "SPV"],
      ["QA & QC Staff", "Staff"], ["Project & Design Supervisor", "SPV"], ["Admin Project", "Staff"],
      ["Assistant Manager HRD", "Manager"], ["HRD Supervisor", "SPV"], ["HRD Staff", "Staff"],
      ["Office Girl", "Staff"], ["Assistant Manager Bussines Development", "Manager"],
      ["Marketing Manager", "Manager"], ["Marketing Supervisor", "SPV"], ["Marketing Staff", "Staff"],
      ["Design Grafis Staff", "Staff"], ["Marketing Senior Staff", "Staff"],
      ["Training Development Manager", "Manager"], ["Training Development Senior Staff", "Staff"],
      ["Training Development Staff", "Staff"], ["Personal Assistant", "Staff"],
      ["Admin Warehouse Senior Staff", "Staff"], ["Admin Logistik", "Staff"], ["Admin Checker", "Staff"],
      ["Admin Gudang", "Staff"], ["Gudang Staff", "Staff"], ["Driver", "Staff"], ["Teknisi Staff", "Staff"],
      ["Store Supervisor", "SPV"], ["Cashier Leader", "SPV"], ["Leader kitchen", "SPV"],
      ["Cook Helper Agemono", "Staff"], ["Cook Helper Sushi", "Staff"], ["Cashier", "Staff"], ["Stocker", "Staff"]
    ];
    shLevel.getRange(2, 1, levelSeed.length, 2).setValues(levelSeed);
  }

  // 2. Tambah + isi kolom "Tipe Lokasi" di MasterLokasiKerja kalau belum ada
  const shLokasi = ss.getSheetByName("MasterLokasiKerja");
  if (shLokasi) {
    const headers = shLokasi.getRange(1, 1, 1, shLokasi.getLastColumn()).getValues()[0];
    if (headers.indexOf("Tipe Lokasi") === -1) {
      const newCol = shLokasi.getLastColumn() + 1;
      shLokasi.getRange(1, newCol).setValue("Tipe Lokasi")
        .setFontWeight("bold").setBackground("#008bca").setFontColor("#ffffff");

      const lastRow = shLokasi.getLastRow();
      if (lastRow > 1) {
        const kodeLokasiVals = shLokasi.getRange(2, 1, lastRow - 1, 1).getValues();
        const tipeVals = kodeLokasiVals.map(r => {
          const kode = String(r[0]).trim().toUpperCase();
          return [(kode === "MMJ-HOF" || kode === "MMJ-GMR") ? "Head Office" : "Store"];
        });
        shLokasi.getRange(2, newCol, tipeVals.length, 1).setValues(tipeVals);
      }
      Logger.log("Kolom 'Tipe Lokasi' berhasil ditambahkan & diisi otomatis di MasterLokasiKerja.");
    }
  }

  Logger.log("setupLevelJabatanDanTipeLokasi selesai.");
}

/**
 * Setup modul Absensi (integrasi ADMS/iClock). Membuat sheet AbsensiHarian untuk menyimpan
 * hasil sync dari database MySQL ADMS (checkinout + userinfo), dikelompokkan per karyawan
 * per tanggal. Aman dijalankan berkali-kali.
 */
function setupAbsensiModule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let shAbsensi = ss.getSheetByName("AbsensiHarian");
  if (!shAbsensi) {
    shAbsensi = ss.insertSheet("AbsensiHarian");
    shAbsensi.appendRow([
      "ID Karyawan", "Nama", "Kode Lokasi", "Tanggal",
      "Jam Masuk", "Jam Pulang", "Status", "Terlambat (Menit)", "Tanggal Sync", "Mesin Masuk", "Mesin Pulang"
    ]);
    shAbsensi.getRange(1, 1, 1, 11).setFontWeight("bold").setBackground("#008bca").setFontColor("#ffffff");
    shAbsensi.setFrozenRows(1);
  }

  Logger.log("Modul Absensi (ADMS) berhasil diinisialisasi.");
}

function setupPayrollModule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. PayrollSettings - Konfigurasi global (nominal yang sama untuk semua karyawan)
  let shSettings = ss.getSheetByName("PayrollSettings");
  if (!shSettings) {
    shSettings = ss.insertSheet("PayrollSettings");
    shSettings.appendRow(["Key", "Value", "Keterangan"]);
    shSettings.appendRow(["AbsensiJedaSesiJam", 14, "Jeda antar scan (jam) yang memulai sesi absensi baru"]);
    shSettings.appendRow(["AbsensiMaxDurasiJam", 16, "Durasi maksimal 1 sesi absensi (jam)"]);
    shSettings.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#008bca").setFontColor("#ffffff");
    shSettings.setFrozenRows(1);
  }

  // 2. MasterLokasiKerja - Jadwal kerja & pembagi prorate per lokasi
  let shLokasi = ss.getSheetByName("MasterLokasiKerja");
  if (!shLokasi) {
    shLokasi = ss.insertSheet("MasterLokasiKerja");
    shLokasi.appendRow(["Kode Lokasi", "Nama Lokasi", "Jadwal Kerja", "Jam Kerja", "Hari Off", "Pembagi Hari Kerja", "Tipe Lokasi"]);
    const defaultLokasi = [
      ["MMJ-HOF", "Head Office", "Senin-Jumat", "09:00-18:00", "Sabtu-Minggu", 21, "Head Office"],
      ["MMJ-GMR", "Gudang Marlina", "Senin-Sabtu", "Senin-Jumat 09:00-17:00 / Sabtu 09:00-14:00", "Minggu", 25, "Head Office"],
      ["MMJ-GGP", "Grand Galaxy Park", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-LWB", "Living World Bekasi", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-LMC", "Lippo Mall Cikarang", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-RPK", "Resinda Park Karawang", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-MCD", "Margo City Depok", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-BSQ", "Botani Square", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-PIK", "Pantai Indah Kapuk", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-MOI", "Mall Of Indonesia", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-GI", "Grand Indonesia", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-ECP", "Express Capital Place", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-BNI 46", "Wisma BNI 46", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-BMP", "Blok M Plaza", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-GMP", "Gajah Mada Plaza", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-MCJ", "Cijantung Mall", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-TPS", "The Park Sawangan", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-LWS", "Living World Alam Sutera", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-BTR", "Bintaro Plaza", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-TGC", "Tangerang City", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-SKC", "Supermall Karawaci", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-TSC", "Trans Studio Cibubur", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-GMC", "Grage Mall Cirebon", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-SSQ", "Solo Square", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-SPM", "Solo Paragon", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-PMJ", "Pakuwon Mall Yogyakarta", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-APJ", "Ambarukmo Plaza", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-JCM", "Jogja City Mall", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-PRS", "Paragon Mall Semarang", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-QCS", "Queen City Semarang", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-RPS", "Royal Plaza Surabaya", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-CWS", "Ciputra World Surabaya", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-LCM", "Lampung City Mall", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-PSQ", "Palembang Square", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-CPB", "Ciputra Pekanbaru", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-PMF", "Plaza Medan Fair", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-MTS", "Manhattan Medan", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-TSB", "Trans Studio Bandung", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-BIP", "Bandung Indah Plaza", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-DBB", "Dbotanica Bandung", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-FCL", "Festival Citylink Bandung", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-TMP", "Transmart Padang", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-WTJ", "WTC Jambi", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-LWD", "Living World Denpasar Bali", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-BWB", "Beachwalk Bali", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-MOG", "Mall Olympic Garden Malang", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-MSQ", "Malang Town Square", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-CCM", "Cibinong City Mall", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-KKB", "Kota Kasablanka", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-AYP", "Ayani Mall Pontianak", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-GBR", "Gaia Bumi Raya", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-ITC BSD", "ITC BSD", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"],
      ["MMJ-CPM", "Centre Point Medan", "Shifting/Senin-Sabtu", "", "Minggu", 25, "Store"]
    ];
    shLokasi.getRange(2, 1, defaultLokasi.length, 7).setValues(defaultLokasi);
    shLokasi.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#008bca").setFontColor("#ffffff");
    shLokasi.setFrozenRows(1);
    // Catatan: Jam Kerja outlet di atas sengaja dikosongkan (selain Head Office/Gudang Marlina) -
    // silakan lengkapi per outlet lewat menu Master Lokasi Kerja. Pembagi Hari Kerja default 25,
    // koreksi manual untuk outlet dengan jadwal 5 hari kerja. Tipe Lokasi menentukan aturan
    // komponen gaji bulanan/harian per level jabatan (lihat MasterLevelJabatan).
  }

  // 2b. MasterLevelJabatan - Klasifikasi level (Staff/SPV/Manager) per Jabatan, dipakai untuk
  // menentukan otomatis apakah komponen gaji karyawan itu Bulanan atau Harian:
  //   Staff          -> selalu Harian (semua lokasi)
  //   SPV            -> Bulanan di Head Office, Harian di Store
  //   Manager ke atas -> selalu Bulanan (semua lokasi)
  let shLevel = ss.getSheetByName("MasterLevelJabatan");
  if (!shLevel) {
    shLevel = ss.insertSheet("MasterLevelJabatan");
    shLevel.appendRow(["Jabatan", "Level"]);
    const levelSeed = [
      ["Direktur", "Manager"], ["Logistik & Warehouse Manager", "Manager"], ["Area Manager", "Manager"],
      ["Chef", "SPV"], ["Trainer", "SPV"],
      ["Internal Audit Supervisor", "SPV"], ["Internal Audit Staff", "Staff"],
      ["Manager Finance & Accounting", "Manager"], ["Assistant Manager Finance & Accounting", "Manager"],
      ["Finance Supervisor", "SPV"], ["Accounting Supervisor", "SPV"],
      ["Account Payable Staff", "Staff"], ["Account Receivable Staff", "Staff"],
      ["Finance Accountig Tax Spv", "SPV"], ["Accounting Staff", "Staff"], ["Finance Staff", "Staff"],
      ["Kurir", "Staff"],
      ["Purchasing Manager", "Manager"], ["Purchasing Supervisor", "SPV"],
      ["Purchasing Senior Staff", "Staff"], ["Purchasing Staff", "Staff"],
      ["IT Support", "Staff"], ["Driver Pribadi", "Staff"],
      ["QA & QC Supervisor", "SPV"], ["QA & QC Staff", "Staff"],
      ["Project & Design Supervisor", "SPV"], ["Admin Project", "Staff"],
      ["Assistant Manager HRD", "Manager"], ["HRD Supervisor", "SPV"], ["HRD Staff", "Staff"],
      ["Office Girl", "Staff"],
      ["Assistant Manager Bussines Development", "Manager"], ["Marketing Manager", "Manager"],
      ["Marketing Supervisor", "SPV"], ["Marketing Staff", "Staff"],
      ["Design Grafis Staff", "Staff"], ["Marketing Senior Staff", "Staff"],
      ["Training Development Manager", "Manager"], ["Training Development Senior Staff", "Staff"],
      ["Training Development Staff", "Staff"], ["Personal Assistant", "Staff"],
      ["Admin Warehouse Senior Staff", "Staff"], ["Admin Logistik", "Staff"], ["Admin Checker", "Staff"],
      ["Admin Gudang", "Staff"], ["Gudang Staff", "Staff"], ["Driver", "Staff"], ["Teknisi Staff", "Staff"],
      ["Store Supervisor", "SPV"], ["Cashier Leader", "SPV"], ["Leader kitchen", "SPV"],
      ["Cook Helper Agemono", "Staff"], ["Cook Helper Sushi", "Staff"], ["Cashier", "Staff"], ["Stocker", "Staff"]
    ];
    shLevel.getRange(2, 1, levelSeed.length, 2).setValues(levelSeed);
    shLevel.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#008bca").setFontColor("#ffffff");
    shLevel.setFrozenRows(1);
  }

  // 3. MasterGajiKaryawan - Komponen gaji tetap per karyawan
  let shGaji = ss.getSheetByName("MasterGajiKaryawan");
  if (!shGaji) {
    shGaji = ss.insertSheet("MasterGajiKaryawan");
    shGaji.appendRow([
      "ID Karyawan", "Nama", "Kode Lokasi", "Jabatan", "Status Karyawan",
      "Gaji Pokok", "Tunjangan Jabatan", "Tunjangan Makan Bulanan", "Tunjangan Transport Bulanan",
      "Tunjangan Makan Harian", "Tunjangan Transport Harian",
      "No Rekening", "Nama Bank", "NPWP", "Status PTKP", "Tanggal Masuk", "Status Aktif",
      "Jatah Cuti Tahunan"
    ]);
    shGaji.getRange(1, 1, 1, 18).setFontWeight("bold").setBackground("#008bca").setFontColor("#ffffff");
    shGaji.setFrozenRows(1);
  }

  // 4. PayrollInput - Input manual absensi & komponen variabel per periode
  let shInput = ss.getSheetByName("PayrollInput");
  if (!shInput) {
    shInput = ss.insertSheet("PayrollInput");
    shInput.appendRow([
      "ID Input", "Bulan", "Tahun", "ID Karyawan",
      "Hari Kerja Aktual", "Jumlah Alpa", "Jumlah Sakit Tanpa Surat", "Jumlah Izin",
      "Jumlah Cuti", "Jumlah Sakit Dengan Surat", "Datang Terlambat (Kali)", "Pulang Cepat (Kali)",
      "Uang Kerajinan", "Tunjangan Toko", "Tunjangan Multiskill", "Pendapatan Lainnya", "Rapel Manual",
      "Potongan Varian", "Deposit Seragam", "Potongan Deposit Training", "Potongan Karyawan", "Potongan Pph21",
      "Catatan", "Input By", "Tanggal Input"
    ]);
    shInput.getRange(1, 1, 1, 25).setFontWeight("bold").setBackground("#008bca").setFontColor("#ffffff");
    shInput.setFrozenRows(1);
  }

  Logger.log("Modul Penggajian berhasil diinisialisasi.");
}


/**
 * Setup modul Hak Akses (Permissions) - per Role & per User.
 * Jalankan SEKALI dari editor Apps Script (pilih fungsi ini lalu klik Run).
 * Aman dijalankan berkali-kali: hanya mengisi default jika sheet masih kosong,
 * tidak akan menimpa perubahan hak akses yang sudah pernah disimpan dari UI.
 */
function setupPermissionsModule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Permissions");

  if (!sheet) {
    sheet = ss.insertSheet("Permissions");
    sheet.appendRow(["ID", "ScopeType", "ScopeValue", "ModuleKey", "FeatureKey", "Allowed"]);
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 6, 140);
  }

  if (sheet.getLastRow() < 2) {
    const roles = ["Admin", "HR", "Manager", "MOD"];
    const rows = [];
    let n = 0;

    MODULE_DEFINITIONS_.forEach(function (mod) {
      const featureKeys = ["view"].concat(mod.features || []);
      roles.forEach(function (role) {
        featureKeys.forEach(function (feat) {
          n++;
          rows.push([
            "PERM-" + String(n).padStart(4, "0"),
            "ROLE",
            role,
            mod.key,
            feat,
            defaultAllowedForRole_(role, mod.key, feat)
          ]);
        });
      });
    });

    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }

  Logger.log("Setup modul Permissions selesai. Total baris: " + (sheet.getLastRow() - 1));
}