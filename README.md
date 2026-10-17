# Aplikasi SPD Puskesmas Moncongloe

Aplikasi web (PWA) berbentuk dashboard untuk membuat empat dokumen perjalanan dinas dari satu isian kegiatan:

1. **SPD** (Excel): satu sheet per pegawai, A4 portrait, dua halaman timbal balik.
2. **Surat Tugas** (Word).
3. **Lembar Penerimaan** uang transport / LPJ (Excel).
4. **Lembar Verifikasi** (Excel).

Kegiatan yang sudah dibuat **tersimpan**, jadi revisi cukup dibuka lagi dari menu **Daftar Kegiatan**
tanpa mengunggah template. Data disimpan di Google Sheet "Database SPD Moncongloe" lewat Apps Script
(lihat [apps-script/PANDUAN.md](apps-script/PANDUAN.md)). Selama Google Sheet belum dihubungkan,
data disimpan di browser perangkat itu saja.

## Menu

| Menu | Isi |
|---|---|
| Dashboard | Ringkasan jumlah kegiatan, kunjungan, pegawai, dan kegiatan terbaru. |
| Buat Kegiatan Baru | Form Surat Tugas + tabel Penugasan (satu baris = satu kunjungan). Bisa juga **Impor dari Excel** memakai template lama. |
| Daftar Kegiatan | Semua kegiatan tersimpan: cari, buka untuk revisi, unduh ulang, hapus. |
| Data Pegawai | Ubah, tambah, hapus pegawai. Dipakai untuk pilihan nama di Penugasan. |
| Data Posyandu | Daftar posyandu dan desanya. Tempat Tujuan cukup diketik nama posyandunya (mis. `tokka`), lalu otomatis menjadi `Tokka Desa Bonto Marannu`, sehingga kolom Desa di Surat Tugas terisi. |
| Pengaturan | Koneksi Google Sheet (URL Web App + PIN) dan nilai bawaan form (pejabat, PPTK, anggaran, verifikator). |

## Alur pakai

1. **Buat Kegiatan Baru**: isi Nomor ST (lengkap), Tanggal Surat Tugas, nama kegiatan, lalu baris penugasan
   (No SPD, pegawai, tempat tujuan, tanggal). Tombol **Beri nomor otomatis** membagikan Nomor SPD
   berurutan, satu nomor per pegawai.
2. **Periksa**. Pesan merah wajib diperbaiki sebelum mengunduh; pesan kuning boleh dilewati.
3. **Simpan**. Kegiatan yang belum lengkap tetap bisa disimpan sebagai draf.
4. **Unduh ▴** lalu pilih salah satu dokumen atau **Semua (4 file)**.

## Tanggal Surat Tugas dan Tanggal SPD

Satu kegiatan = satu Surat Tugas dengan satu nomor dan satu tanggal. Tanggal SPD boleh berbeda:
*Tanggal SPD* di bagian atas berlaku untuk semua SPD (kosong = sama dengan Tanggal Surat Tugas), dan
kolom *Tgl SPD* di Penugasan dipakai bila SPD pegawai tertentu bertanggal lain. Cukup isi di salah satu
baris pegawai itu. Bulan romawi dan tahun di Nomor SPD mengikuti tanggal SPD-nya.

## Tanpa internet

Setelah pernah dibuka sekali dalam keadaan online, aplikasi tetap bisa dipakai tanpa internet:
daftar kegiatan, pegawai, dan posyandu memakai salinan terakhir dari Google Sheet, dan keempat file tetap
bisa diunduh. Kegiatan yang disimpan saat offline diberi tanda **Belum tersinkron** dan dikirim otomatis
ke Google Sheet begitu internet tersambung.

Bila dua petugas mengubah kegiatan yang sama, aplikasi tidak menimpa diam-diam. Saat online, petugas
ditanya apakah mau menimpa. Saat sinkron dari offline, versi perangkat disimpan sebagai kegiatan
terpisah berakhiran "(salinan offline)".

## Aturan Lembar Penerimaan dan Lembar Verifikasi

| Hal | Aturan |
|---|---|
| Lembar Penerimaan | Satu baris per kunjungan, urut Nomor SPD lalu tanggal. Uang transport per kunjungan dari isian *Uang transport* (bawaan Rp 100.000). Jumlah memakai rumus SUM dan terbilang otomatis. Kop, Program, Kegiatan, Sub Kegiatan, PPTK, dan pejabat dari bagian *Anggaran* dan *Penanda tangan*. |
| Lembar Verifikasi | Satu baris per tanggal + tempat (sama dengan jadwal Surat Tugas), berisi hari, tempat, dan nama petugas. Paling sedikit 8 baris seperti cetakan aslinya. Nama verifikator dari bagian *Verifikator*. |

Cetakannya `assets/master_penerimaan.xlsx` dan `assets/master_verifikasi.xlsx`. Posisi sel yang diisi
diatur di fungsi `buatPenerimaan` dan `buatVerifikasi` dalam `js/spd-core.js`.

## Aturan pembuatan Surat Tugas

Satu template = satu Surat Tugas (.docx) berisi semua pegawai, dari cetakan `assets/master_st.docx`.

| Hal | Aturan |
|---|---|
| Nomor, tanggal, pejabat | Dari sheet Kegiatan: *Nomor Surat Tugas*, *Tanggal dikeluarkan*, *Nama/NIP Pejabat*, *Pengguna Anggaran* (jabatan penanda tangan). |
| Dasar | Baris *Dasar Surat Tugas* di sheet Kegiatan. |
| Untuk | `Melaksanakan Kegiatan` + *Maksud Perjalanan Dinas*. |
| Tabel jadwal | Satu baris per kombinasi tanggal + tempat (tanpa duplikat), urut tanggal, nama hari otomatis. *Tempat Tujuan* dipisah di kata `Desa`: `Tokka Desa Bonto Marannu` → Posyandu `Tokka`, Desa `Bonto Marannu`. |
| Lampiran | Satu baris per pegawai, urut Nomor SPD, berisi semua tempat kegiatannya. |

Kop, logo, dan teks baku Surat Tugas bisa diubah langsung di `assets/master_st.docx` dengan Word, asalkan
penanda seperti `{{NOMOR_ST}}`, `{{DASAR}}`, `{{J_NO}}`, `{{L_NAMA}}` tidak dihapus.

## Aturan pembuatan SPD

| Hal | Aturan |
|---|---|
| Nomor SPD | Petugas cukup mengetik angka urut (mis. `2006`). Nomor lengkap dirakit otomatis: `2006/` + *Kode Nomor SPD* (sheet Kegiatan, mis. `SPD/PKM-ML`) + `/` + bulan romawi + `/` + tahun dari *Tanggal dikeluarkan (SPD)* → `2006/SPD/PKM-ML/IX/2026`. Nomor lengkap yang diketik manual tetap diterima. Satu pegawai = satu nomor. |
| Lebih dari 6 kunjungan | Tetap satu SPD dengan nomor yang sama. Halaman depan memuat semua tanggal & tujuan; halaman belakang ditambah (6 kunjungan per halaman). Slot I diisi di tiap halaman belakang, bagian VIII *Tiba kembali* hanya di halaman belakang terakhir. |
| Nama sheet | `NamaDepan_dd-mm`, dengan tanggal kunjungan pertama (contoh `Riski_06-08`). |
| Lama perjalanan | Jumlah kunjungan, misalnya `3 (Tiga) Hari`. Satu kunjungan = satu hari, tanpa menginap. |
| Tanggal berangkat | `06, 13, 15 Agustus 2026`. Lintas bulan: `28 Juli, 02 Agustus 2026`. |
| Pangkat/Jabatan | Dicetak apa adanya dari sheet Pegawai. |
| Urutan sheet | Menurut Nomor SPD. |

## Memasang di GitHub Pages (sekali saja)

1. Buat repository baru di GitHub, misalnya `spd-moncongloe`. Repo boleh **Public**.
   Isinya hanya kode, template, dan daftar pegawai, tanpa data SPD.
2. Unggah **seluruh isi** folder ini ke repo: **Add file → Upload files**, lalu tarik semua file
   dan folder, termasuk file `.nojekyll`.
3. Buka **Settings → Pages**. Di bagian *Source*, pilih **Deploy from a branch**, branch `main`,
   folder `/ (root)`, lalu **Save**.
4. Tunggu 1–2 menit. Aplikasi tersedia di `https://<username>.github.io/spd-moncongloe/`.

> Catatan: sheet Pegawai memuat NIP/NIK. Kalau ingin repo **Private**, GitHub Pages dari repo
> private memerlukan akun GitHub berbayar (Pro/Team).

## Memperbarui daftar pegawai

Setelah Google Sheet terhubung, cukup ubah lewat menu **Data Pegawai**. Cara di bawah hanya untuk
template Excel yang dipakai fitur Impor dan sebagai daftar awal sebelum Google Sheet terhubung.

1. Unduh `assets/Template_Input_SPD.xlsx` dari repo.
2. Ubah sheet **Pegawai** di Excel (tambah baris, ubah pangkat, dan lain-lain), lalu simpan.
   NIP/NIK harus berupa **teks**: awali dengan tanda petik `'` bila Excel mengubahnya menjadi angka.
3. Di GitHub, buka folder `assets` → **Add file → Upload files**, unggah file dengan nama yang
   **sama persis** (`Template_Input_SPD.xlsx`), lalu **Commit**.

Petugas otomatis mendapat versi baru saat membuka aplikasi dalam keadaan online. Kalau petugas
mengunggah template lama yang daftar pegawainya berbeda, aplikasi memberi peringatan.

Tidak perlu mengubah kode atau `sw.js` saat memperbarui template.

## Mengubah tata letak SPD (kop, logo, teks baku)

Cetakan SPD ada di `assets/master_spd.xlsx`. Sel yang diisi aplikasi diatur di fungsi `isiSel`
dalam `js/spd-core.js`. Kalau posisi baris/kolom di cetakan diubah, fungsi tersebut harus ikut
disesuaikan.

## Struktur folder

```
index.html              halaman aplikasi (dashboard)
css/app.css             tampilan
js/app.js               antarmuka: menu, form kegiatan, data pegawai, pengaturan
js/config.js            URL Web App Google Sheet untuk semua petugas
js/simpan.js            penyimpanan data (Google Sheet atau browser)
js/spd-core.js          pemeriksa dan pembuat SPD, Surat Tugas, Lembar Penerimaan, Lembar Verifikasi
js/vendor/              SheetJS (baca Excel) & JSZip (tulis Excel)
assets/master_spd.xlsx  cetakan SPD
assets/master_st.docx   cetakan Surat Tugas
assets/master_penerimaan.xlsx   cetakan Lembar Penerimaan (LPJ)
assets/master_verifikasi.xlsx   cetakan Lembar Verifikasi
assets/posyandu.json    daftar posyandu awal (sebelum Google Sheet terhubung)
icons/logo-maros.png    logo di sidebar
apps-script/            Code.gs, template Google Sheet, dan panduan pemasangannya
assets/Template_Input_SPD.xlsx   template kosong untuk petugas
contoh/                 contoh template terisi (ST Kelas Ibu Balita Agustus 2026)
sw.js, manifest.webmanifest, icons/   bagian PWA (offline & pasang)
```

Developed by Manjilala-Poltekkes Kemenkes Makassar/Persagi Sulsel.
