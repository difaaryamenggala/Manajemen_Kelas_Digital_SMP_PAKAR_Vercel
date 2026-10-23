# Manajemen Kelas Digital SMP PAKAR — Vercel + Google Apps Script

## Isi paket
- `index.html`: antarmuka aplikasi, bridge API Vercel, pemindai QR dan kamera selfie depan.
- `api/gs.js`: Vercel Serverless Function sebagai proxy ke Google Apps Script.
- `Code.gs`: backend Google Apps Script dan fungsi database.
- `vercel.json`, `package.json`: konfigurasi deploy.

## Penting
Isi video YouTube belum dapat diverifikasi dari tautannya saja. Paket ini mengadaptasi kode yang Anda unggah, bukan klaim salinan persis tutorial video.

## Langkah 1 — Siapkan Google Sheets dan Apps Script
1. Buat Google Spreadsheet baru.
2. Buka **Extensions > Apps Script**.
3. Salin seluruh isi `Code.gs` ke file script utama (misalnya `Code.gs`). Jangan tempel file `setup(1).gs` kedua kali karena `setupSheet()` sudah ada di `Code.gs`.
4. Pastikan project Apps Script terikat pada spreadsheet yang benar.
5. Jalankan fungsi `setupSheet()` dari editor Apps Script, lalu izinkan permission Google Sheets dan Google Drive.
6. Deploy > **New deployment** > pilih **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone** (dibutuhkan oleh proxy Vercel; baca catatan keamanan di bawah)
7. Deploy dan salin URL yang berakhir `/exec`.

## Langkah 2 — Upload ke Vercel
1. Unggah folder proyek ini ke repository GitHub, atau jalankan `vercel` dari folder proyek menggunakan Vercel CLI.
2. Di Vercel, buka **Project > Settings > Environment Variables**.
3. Tambahkan:
   - Name: `GAS_WEB_APP_URL`
   - Value: URL Web App Apps Script yang berakhir `/exec`
   - Environment: Production (tambahkan Preview juga jika diperlukan)
4. Deploy / Redeploy proyek.
5. Buka domain HTTPS Vercel. Kamera hanya dapat digunakan di secure context (HTTPS atau localhost), dan browser harus diberi izin kamera.

## Langkah 3 — Kamera
- Tombol **Buka Kamera Presensi** memakai kamera belakang untuk memindai QR.
- Setelah QR terbaca, aplikasi meminta izin kamera depan dan menampilkan layar **Ambil Foto Selfie Presensi**.
- Tekan **Ambil & Kirim** untuk mengirim QR dan foto ke Apps Script.
- Jika kamera tidak muncul, periksa izin kamera situs di browser, tutup aplikasi lain yang memakai kamera, dan pastikan membuka domain HTTPS.

## Catatan keamanan penting sebelum digunakan dengan data siswa sungguhan
Kode dasar yang diunggah memakai UserID/password biasa dan sejumlah fungsi menerima objek `currentUser` dari browser. Itu tidak cukup aman untuk data siswa dalam produksi karena permintaan browser bisa dimanipulasi. Endpoint Apps Script juga perlu otorisasi berbasis sesi/token server-side, pembatasan role pada setiap fungsi, serta kebijakan foto yang sesuai. `Anyone` pada deployment Apps Script berarti endpoint dapat dipanggil dari luar aplikasi; URL bukanlah sistem autentikasi. Uji dahulu dengan data contoh dan jangan gunakan data sensitif sampai pengamanan backend selesai.

## Tes cepat
- Login menggunakan akun yang memang sudah ada pada sheet `Users`.
- Pastikan tab `Users` memiliki header `UserID, Password, Nama, Role, Kelas, Email, NoHP`.
- Tes kamera di Chrome desktop/Android atau Safari iOS dengan izin kamera aktif.
- Uji QR yang valid, QR kedaluwarsa, dan presensi ganda.

## Batasan integrasi
Paket ini menghubungkan antarmuka Vercel ke Apps Script lewat endpoint `/api/gs`. Jangan buka `index.html` langsung sebagai `file://` jika ingin kamera bekerja. Jalankan melalui domain HTTPS Vercel.
