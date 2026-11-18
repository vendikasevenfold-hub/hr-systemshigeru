# Memasang Shigeru HR di VPS Windows Server (bersama GLPI)

Di VPS, Shigeru HR dipasang **langsung tanpa Docker**: Node.js + PostgreSQL untuk Windows, berjalan
sebagai Windows Service. Docker di Windows Server tidak bisa menjalankan container Linux, sedangkan
image aplikasi ini memakai Linux. Apache yang sudah dipakai GLPI meneruskan subdomain HR ke aplikasi.

```
Internet ──https://hr.contoh.co.id──▶ Apache (port 443, dipakai bersama GLPI)
                                         │  ProxyPass
                                         ▼
                               Shigeru HR (Node.js) 127.0.0.1:3100
                                         │
                                         ▼
                               PostgreSQL 127.0.0.1:5432   (GLPI tetap pakai MySQL/MariaDB)
```

Ganti `hr.contoh.co.id` dengan subdomain Anda di seluruh panduan ini.
Semua perintah dijalankan di VPS lewat Remote Desktop, di **PowerShell "Run as Administrator"**,
kecuali yang ditandai *di laptop*.

---

## 0. Sebelum mulai

- [ ] **DNS:** buat record `A` untuk `hr.contoh.co.id` ke IP publik VPS (di panel domain/DNS).
- [ ] **Zona waktu VPS:** Settings → Time & language → Time zone = **(UTC+07:00) Bangkok, Hanoi, Jakarta**.
- [ ] **Backup GLPI** dulu (folder GLPI + database MySQL), untuk berjaga-jaga sebelum Apache diubah.
- [ ] Catat lokasi Apache: `C:\xampp\apache` (XAMPP), `C:\Apache24` (Apache Lounge), atau lainnya.

## 1. Pasang Node.js

Unduh **Node.js 22 LTS (Windows Installer .msi, x64)** dari https://nodejs.org lalu pasang dengan
pilihan bawaan. Cek:

```powershell
node --version   # v22.x
```

## 2. Pasang PostgreSQL

Unduh **PostgreSQL 16 untuk Windows** dari https://www.postgresql.org/download/windows/ (installer EDB).

- Password untuk user `postgres`: buat yang kuat dan simpan baik-baik.
- Port: `5432`. Locale: bawaan.
- Stack Builder di akhir instalasi tidak perlu.

Buat database dan user aplikasi (ganti `PASSWORD_DB` dengan password baru, huruf dan angka saja):

```powershell
$psql = "C:\Program Files\PostgreSQL\16\bin\psql.exe"
& $psql -U postgres -c "CREATE ROLE shigeru_hr LOGIN PASSWORD 'PASSWORD_DB';"
& $psql -U postgres -c "CREATE DATABASE shigeru_hr OWNER shigeru_hr;"
```

PostgreSQL bawaan hanya menerima koneksi dari VPS itu sendiri. Port 5432 **jangan** dibuka di firewall.

## 3. Ambil kode aplikasi

Pilih salah satu:

- **Dengan Git** (disarankan, supaya mudah update). Pasang Git for Windows dari https://git-scm.com, lalu:
  ```powershell
  git clone https://github.com/vendikasevenfold-hub/hr-systemshigeru C:\shigeru-hr
  ```
- **Tanpa Git:** di GitHub klik **Code → Download ZIP**, ekstrak ke `C:\shigeru-hr`.

Pasang dependensi:

```powershell
cd C:\shigeru-hr
npm ci --omit=dev
```

## 4. Isi konfigurasi `.env`

```powershell
Copy-Item C:\shigeru-hr\deploy\windows\.env.vps.example C:\shigeru-hr\.env
notepad C:\shigeru-hr\.env
```

Ganti `GANTI_PASSWORD_DB` dengan password dari langkah 2, dan isi `ADMS_BRIDGE_URL` / `ADMS_BRIDGE_KEY`.

## 5. Pindahkan data dari laptop

*Di laptop* (Docker sedang jalan), dari folder repo:

```powershell
powershell -ExecutionPolicy Bypass -File deploy\export-db-dari-laptop.ps1
```

Salin `shigeru-hr.dump` ke VPS (copy-paste lewat Remote Desktop), misalnya ke `C:\shigeru-hr-backup\`.
Lalu *di VPS*:

```powershell
$env:PGPASSWORD = "PASSWORD_DB"
& "C:\Program Files\PostgreSQL\16\bin\pg_restore.exe" -h 127.0.0.1 -U shigeru_hr -d shigeru_hr --no-owner --no-privileges C:\shigeru-hr-backup\shigeru-hr.dump
Remove-Item Env:PGPASSWORD
```

Hapus `shigeru-hr.dump` dari laptop setelah berhasil (isinya data gaji karyawan).

## 6. Jalankan sebagai Windows Service

Unduh **NSSM** dari https://nssm.cc/download, ekstrak, lalu salin `win64\nssm.exe` ke `C:\nssm\nssm.exe`.

```powershell
powershell -ExecutionPolicy Bypass -File C:\shigeru-hr\deploy\windows\install-service.ps1 -NssmPath C:\nssm\nssm.exe
```

Harus muncul `OK: aplikasi menjawab di http://127.0.0.1:3100`. Tes juga di browser VPS:
http://127.0.0.1:3100 lalu login. Log ada di `C:\shigeru-hr\logs\shigeru-hr.log`.

## 7. Apache & GLPI

> **Penting.** Kalau GLPI sekarang **tidak** memakai `<VirtualHost>` (langsung dari `DocumentRoot`
> di httpd.conf), maka begitu Anda menambah VirtualHost untuk HR, Apache akan menyajikan **HR untuk
> semua alamat**, termasuk alamat GLPI. Cek dulu:
>
> ```powershell
> Select-String -Path C:\xampp\apache\conf\httpd.conf, C:\xampp\apache\conf\extra\*.conf -Pattern "<VirtualHost" 
> ```
>
> Kalau tidak ada VirtualHost untuk GLPI, tambahkan dulu satu VirtualHost untuk GLPI **di atas**
> blok HR (ganti nama domain dan folder GLPI sesuai server Anda):
>
> ```apache
> <VirtualHost *:80>
>     ServerName glpi.contoh.co.id
>     DocumentRoot "C:/xampp/htdocs/glpi/public"
> </VirtualHost>
> ```
>
> Kalau GLPI sudah memakai HTTPS (port 443), salin juga blok `<VirtualHost *:443>` GLPI yang lama.

Lalu:

1. Salin `deploy\windows\apache-shigeru-hr.conf` ke `...\apache\conf\extra\shigeru-hr.conf`,
   ganti `hr.contoh.co.id` dengan subdomain Anda.
2. Di `httpd.conf`, aktifkan modul yang disebut di bagian atas file itu, lalu tambahkan
   `Include conf/extra/shigeru-hr.conf` di paling bawah.
3. Untuk sementara (sebelum sertifikat ada), beri tanda `#` pada seluruh blok `<VirtualHost *:443>`
   di `shigeru-hr.conf`.
4. Cek konfigurasi lalu restart Apache:
   ```powershell
   C:\xampp\apache\bin\httpd.exe -t        # harus "Syntax OK"
   ```
   Restart lewat XAMPP Control Panel, atau `Restart-Service Apache2.4` (nama service bisa berbeda).
5. Pastikan **GLPI masih bisa dibuka** seperti biasa.

## 8. HTTPS (sertifikat gratis Let's Encrypt)

Kalau GLPI sudah punya cara memasang sertifikat, pakai cara yang sama. Kalau belum, pakai **win-acme**:

1. Unduh win-acme dari https://www.win-acme.com, ekstrak ke `C:\win-acme`.
2. Buat folder verifikasi: `New-Item -ItemType Directory -Force C:\shigeru-hr-acme\.well-known\acme-challenge`
3. Minta sertifikat:
   ```powershell
   C:\win-acme\wacs.exe --target manual --host hr.contoh.co.id --validation filesystem --webroot C:\shigeru-hr-acme --store pemfiles --pemfilespath C:\shigeru-hr-certs
   ```
   win-acme juga membuat jadwal perpanjangan otomatis.
4. Cek nama file di `C:\shigeru-hr-certs`, sesuaikan tiga baris `SSLCertificate...` di `shigeru-hr.conf`,
   hapus tanda `#` dari blok `<VirtualHost *:443>`, jalankan `httpd.exe -t`, lalu restart Apache.
5. Buka **https://hr.contoh.co.id** dan login.

## 9. Backup otomatis

```powershell
powershell -ExecutionPolicy Bypass -File C:\shigeru-hr\deploy\windows\backup-db.ps1            # tes sekali
powershell -ExecutionPolicy Bypass -File C:\shigeru-hr\deploy\windows\backup-db.ps1 -Register  # tiap hari 01:00
```

Backup disimpan di `C:\shigeru-hr-backup` (14 hari terakhir). Sebaiknya salin juga ke tempat lain
secara berkala (OneDrive kantor / NAS).

## 10. Firewall

Tidak ada port baru yang perlu dibuka. Port 80/443 sudah dibuka untuk GLPI.
Port **3100** (aplikasi) dan **5432** (database) hanya dipakai di dalam VPS. Jangan dibuka.

---

## Update aplikasi nanti

Setelah ada perubahan di GitHub:

```powershell
powershell -ExecutionPolicy Bypass -File C:\shigeru-hr\deploy\windows\update.ps1
```

Skrip ini membuat backup database, mengambil kode terbaru, memasang dependensi, lalu menyalakan
ulang service.

## Kalau ada masalah

| Gejala | Cek |
|---|---|
| Service tidak mau jalan | `C:\shigeru-hr\logs\shigeru-hr.log` |
| Log menampilkan "PERINGATAN: zona waktu" | Zona waktu VPS harus UTC+07:00 (langkah 0) |
| Error koneksi database | Password di `DATABASE_URL` sama dengan langkah 2? Service `postgresql-x64-16` jalan? |
| Subdomain HR menampilkan GLPI, atau GLPI menampilkan HR | Urutan dan `ServerName` VirtualHost di langkah 7 |
| `503 Service Unavailable` dari Apache | Service ShigeruHR berhenti: `Get-Service ShigeruHR`, lalu cek log |
| Semua user harus login ulang | Normal setelah service di-restart (sesi disimpan di memori) |
