# Shigeru HR

Aplikasi HR PT Mideli Makmur Jaya (lembur, data master karyawan, input data bulanan, absensi ADMS,
manajemen user). Dulu berjalan di Google Apps Script + Google Sheets; sekarang berjalan sendiri di
**Docker**: server **Node.js** + database **PostgreSQL**.

## Cara kerja

| Dulu (Apps Script) | Sekarang |
|---|---|
| `Code.gs`, `Setup.gs` | Tetap sama, di `server/gas/`. Dijalankan oleh Node.js. |
| Google Sheets (tab Users, Karyawan, PermintaanLembur, …) | Tabel `sheets` di PostgreSQL. Satu baris = satu tab lama. |
| `google.script.run.namaFungsi()` di `index.html` | Tetap sama. `public/gas-client.js` meneruskannya ke `POST /api/run/namaFungsi`. |
| `SpreadsheetApp`, `LockService`, `CacheService`, `UrlFetchApp`, … | Tiruan di `server/workbook.js` dan `server/gas-runtime.js`. |
| Script Properties `ADMS_BRIDGE_URL` / `ADMS_BRIDGE_KEY` | File `.env`. |

Jadi menambah fitur caranya sama seperti dulu: tulis fungsi di `server/gas/Code.gs`, panggil dari
`public/index.html` dengan `google.script.run`. Fungsi yang namanya diakhiri `_` bersifat privat
(tidak bisa dipanggil dari browser), sama seperti di Apps Script.

## Menjalankan

Syarat: Docker Desktop.

```bash
cp .env.example .env        # lalu isi POSTGRES_PASSWORD (dan ADMS_BRIDGE_* bila perlu)
docker compose up -d --build
```

Buka http://localhost:3100 (atau port `APP_PORT` di `.env`).

Saat database masih kosong, aplikasi membuat tab standar dan user **admin / admin123**.
Segera ganti password-nya lewat menu Manajemen User.

## Memindahkan data dari Google Sheets

1. Buka Google Sheets database HR → **File → Download → Microsoft Excel (.xlsx)**.
2. Simpan filenya di folder `data-import/`, misalnya `data-import/hr.xlsx`.
3. Jalankan:

```bash
docker compose stop app
docker compose run --rm app node scripts/import-xlsx.js /import/hr.xlsx --dry-run
docker compose run --rm app node scripts/import-xlsx.js /import/hr.xlsx
docker compose up -d app
```

Setiap tab di file Excel **menggantikan** tab bernama sama di database. Tab yang tidak ada di file
dibiarkan. `--dry-run` hanya menampilkan ringkasan tanpa menyimpan.

## Pengembangan di VS Code

```bash
npm install
npm test                    # tes fungsi Code.gs tanpa database
docker compose up -d --build   # jalankan ulang setelah mengubah kode
docker compose logs -f app     # lihat log server
```

## Backup database

```bash
docker compose exec db pg_dump -U shigeru_hr shigeru_hr > backup.sql
```

## Catatan

- Server menjalankan satu proses Node.js; jangan menjalankan lebih dari satu container `app`
  terhadap database yang sama.
- Sesi login disimpan di memori, jadi semua user perlu login ulang setelah server di-restart.
- Password user masih disimpan apa adanya (teks biasa), sama seperti di Google Sheets dulu.
