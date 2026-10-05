# Pembuat SPD – UPTD Puskesmas Moncongloe

Aplikasi web (PWA) untuk membuat file **Surat Perjalanan Dinas (SPD)** dari template Excel.
Template masuk, file SPD keluar: satu sheet per SPD, A4 portrait, dua halaman timbal balik
(depan kolom A–G, belakang kolom H–Q; halaman belakang tambahan di kolom R dst. bila kunjungan > 6).

Semua proses berjalan di browser petugas. File **tidak dikirim ke server mana pun**.
Setelah dibuka sekali, aplikasi juga bisa dipakai tanpa internet dan bisa dipasang di HP/laptop.

## Alur pakai

1. Klik **Unduh Template** (selalu versi terbaru dari repo ini).
2. Isi sheet **Kegiatan** (sekali per ST) dan **Penugasan** (satu baris = satu kunjungan).
3. Unggah file ke aplikasi. Aplikasi memeriksa isinya:
   - **Perlu diperbaiki** (merah): proses dihentikan sampai template diperbaiki.
   - **Perhatikan** (kuning): boleh lanjut, tapi sebaiknya dicek.
4. Klik **Buat & Unduh SPD**.

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
index.html              halaman aplikasi
css/app.css             tampilan
js/app.js               antarmuka
js/spd-core.js          pembaca template, pemeriksa, pembuat SPD
js/vendor/              SheetJS (baca Excel) & JSZip (tulis Excel)
assets/master_spd.xlsx  cetakan SPD
assets/Template_Input_SPD.xlsx   template kosong untuk petugas
contoh/                 contoh template terisi (ST Kelas Ibu Balita Agustus 2026)
sw.js, manifest.webmanifest, icons/   bagian PWA (offline & pasang)
```
