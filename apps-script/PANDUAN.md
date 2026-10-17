# Panduan menghubungkan Google Sheet

Langkah ini cukup dilakukan **sekali**, oleh pemilik akun Google yang menyimpan data
(nurfahmi@poltekkes-mks.ac.id). Setelah selesai, semua kegiatan yang dibuat di aplikasi tersimpan di
Google Sheet dan bisa dibuka lagi dari HP atau laptop mana pun.

Yang dibutuhkan dari folder ini:

- `Database_SPD_Moncongloe.xlsx`: template Google Sheet (sheet Pengaturan, Pegawai, Posyandu, Kegiatan, Penugasan, Petunjuk).
- `Code.gs`: kode Apps Script.

## 1. Unggah template ke Google Drive

1. Buka [drive.google.com](https://drive.google.com) dengan akun nurfahmi@poltekkes-mks.ac.id.
2. **Baru → Upload file**, pilih `Database_SPD_Moncongloe.xlsx`.
3. Klik kanan file yang terunggah → **Buka dengan → Google Spreadsheet**.
4. Di Google Sheet: **File → Simpan sebagai Google Spreadsheet**. Pakai file hasil simpanan ini
   (ikon hijau, tanpa label .XLSX). File .xlsx yang asli boleh dihapus.
5. Ganti nama file menjadi **Database SPD Moncongloe** bila perlu.

## 2. Ganti PIN

Buka sheet **Pengaturan**, ganti nilai `PIN` (bawaan `1234`) dengan PIN lain, misalnya 6 angka.
PIN ini diketik petugas di aplikasi. Siapa pun yang tahu PIN dan URL bisa membaca dan mengubah data,
jadi bagikan hanya ke petugas.

Jangan mengubah nama sheet dan baris judul kolom (baris 1) di sheet Pegawai, Posyandu, Kegiatan, dan Penugasan.

## 3. Pasang kode Apps Script

1. Di Google Sheet: **Ekstensi → Apps Script**.
2. Hapus isi file `Kode.gs` yang muncul, lalu tempel **seluruh isi** `Code.gs` dari folder ini.
3. Klik ikon **Simpan** (disket).

## 4. Deploy sebagai Web App

1. Klik **Terapkan (Deploy) → Deployment baru**.
2. Klik ikon roda gigi di samping *Pilih jenis* → **Aplikasi web**.
3. Isi:
   - *Deskripsi*: `SPD Moncongloe`
   - *Jalankan sebagai*: **Saya (nurfahmi@poltekkes-mks.ac.id)**
   - *Yang memiliki akses*: **Siapa saja (Anyone)**
4. Klik **Terapkan**. Google meminta izin: pilih akun, lalu **Izinkan**. Bila muncul layar
   "Google belum memverifikasi aplikasi ini", klik **Lanjutan → Buka SPD Moncongloe (tidak aman)**.
   Ini wajar karena skrip dibuat sendiri.
5. Salin **URL aplikasi web** (berakhiran `/exec`).

Cek: buka URL itu di browser. Bila muncul `"API SPD Moncongloe aktif"`, deploy berhasil.

> **Bila pilihan "Siapa saja" tidak ada** atau URL meminta login akun poltekkes, berarti admin
> Google Workspace kampus membatasi aplikasi web untuk luar domain. Pilihannya: minta admin
> mengizinkan, atau ulangi panduan ini memakai akun Gmail pribadi.

## 5. Hubungkan aplikasi

Pilih salah satu:

- **Untuk semua petugas sekaligus (disarankan)**: buka `js/config.js` di repo GitHub, isi
  `apiUrl: "https://script.google.com/macros/s/…/exec"`, lalu Commit. Petugas tinggal mengetik PIN
  di menu **Pengaturan** aplikasi.
- **Per perangkat**: di aplikasi, buka **Pengaturan**, tempel URL dan PIN, klik **Simpan & Hubungkan**.

Bila sebelumnya sudah ada kegiatan tersimpan di browser, di menu Pengaturan muncul tombol
**Pindahkan data perangkat ke Google Sheet**.

## Bila kode diperbarui

Bila `Code.gs` berubah di kemudian hari: tempel kode baru di Apps Script, lalu
**Terapkan → Kelola deployment → ikon pensil → Versi: Versi baru → Terapkan**. Dengan cara ini
URL `/exec` tetap sama dan `js/config.js` tidak perlu diubah.

## Catatan

- Data pegawai dan posyandu diubah lewat menu **Data Pegawai** dan **Data Posyandu** di aplikasi, atau langsung di sheetnya.
- Nilai bawaan form (pejabat, PPTK, anggaran, verifikator) diubah di menu **Pengaturan** atau sheet Pengaturan.
- Sheet Kegiatan dan Penugasan diisi oleh aplikasi; sebaiknya tidak diubah manual.
- Riwayat perubahan bisa dilihat di **File → Histori versi**, berguna bila ada data terhapus.
