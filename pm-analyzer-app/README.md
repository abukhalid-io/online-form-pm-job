# PM Analyzer App — Scaffold

Starter Next.js untuk input data PM Analyzer (CO/CO2/O2) langsung ke Google Sheets,
dideploy di Vercel. Menggunakan Google Sheets API via service account (bukan Apps Script),
supaya tidak ada masalah CORS dan bisa dikembangkan sebagai aplikasi biasa.

## 1. Setup Google Cloud & Service Account

1. Buka https://console.cloud.google.com, buat project baru (atau pakai yang sudah ada).
2. Aktifkan **Google Sheets API**: menu "APIs & Services" > "Library" > cari "Google Sheets API" > Enable.
3. Buat Service Account: "APIs & Services" > "Credentials" > "Create Credentials" > "Service Account".
   Beri nama bebas, mis. `pm-analyzer-app`.
4. Setelah service account dibuat, buka tab "Keys" > "Add Key" > "Create new key" > pilih **JSON**.
   File JSON akan otomatis terdownload — simpan baik-baik, ini rahasia.
5. Dari file JSON tersebut, kamu butuh 2 nilai:
   - `client_email`
   - `private_key`

## 2. Share Google Sheet ke Service Account

1. Buka Google Sheet PM Analyzer kamu (hasil import file `PM_Analyzer_CO_CO2_O2_Template.xlsx`).
2. Klik tombol **Share**.
3. Tempel `client_email` dari file JSON tadi (bentuknya seperti `xxx@xxx.iam.gserviceaccount.com`).
4. Beri akses **Editor**.
5. Ambil Spreadsheet ID dari URL:
   `https://docs.google.com/spreadsheets/d/`**`INI_ID_NYA`**`/edit`

## 3. Environment Variables

Copy `.env.local.example` jadi `.env.local`, isi 3 nilai di atas. Format `GOOGLE_PRIVATE_KEY`
harus tetap diapit tanda kutip dengan `\n` literal (bukan enter asli) seperti contoh di file.

Saat deploy ke Vercel, isi 3 variabel yang sama di **Project Settings > Environment Variables**.

## 4. Jalankan Lokal

```bash
npm install
npm run dev
```

Buka http://localhost:3000 — form input verifikasi contoh sudah tersambung ke
`PM_Detail_Iterasi` di sheet kamu.

## 5. Struktur yang sudah ada

- `lib/googleSheets.js` — koneksi & helper cari baris berdasar ID
- `pages/api/pm-detail-iterasi.js` — CRUD (GET/POST/PUT/DELETE) untuk sheet `PM_Detail_Iterasi`
- `pages/api/pm-header.js` — CRUD untuk sheet `PM_Header`
- `components/VerificationTimer.jsx` — timer tunggu 5 menit stabilisasi gas, tahan terhadap refresh halaman
- `pages/index.js` — contoh form yang menyatukan semuanya

## 6. Yang masih perlu kamu tambahkan

- API route serupa untuk `Master_Analyzer` dan `Master_Tabung_Gas` (pola sama persis, tinggal
  copy `pages/api/pm-header.js` dan sesuaikan nama sheet & kolom).
- Halaman daftar sesi PM aktif (supaya `id_pm` tidak diketik manual, tapi dipilih dari dropdown).
- Validasi: cegah submit iterasi ke-11 (batas maks sudah divalidasi di sheet lewat dropdown,
  tapi sebaiknya juga dicek di frontend/backend supaya user dapat pesan error yang jelas).
- Auto-hitung `no_iterasi` berikutnya (hitung jumlah baris dengan `id_pm` yang sama, lalu +1),
  supaya user tidak perlu isi manual dan rawan salah nomor.
- Kalau perlu timer yang bisa dipantau lintas device/browser, pindahkan waktu mulai dari
  `localStorage` ke kolom "Waktu" yang sudah ada di `PM_Detail_Iterasi` (baca dari situ saat
  halaman dibuka, bukan dari localStorage).

## Catatan soal kredensial

Service account JSON key **tidak punya masa kadaluarsa otomatis** (beda dengan OAuth token
biasa). Access token yang dipakai tiap request expired tiap 1 jam, tapi library `googleapis`
otomatis refresh sendiri di belakang layar. Key ini berlaku terus sampai kamu revoke manual
dari GCP Console (disarankan rotate tiap beberapa bulan sebagai praktik keamanan, tapi bukan
kewajiban teknis).
