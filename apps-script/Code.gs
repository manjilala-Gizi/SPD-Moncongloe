/*
 * API penyimpanan SPD Moncongloe (Google Apps Script).
 *
 * Tempel seluruh isi file ini di Extensions > Apps Script pada Google Sheet
 * "Database SPD Moncongloe", lalu Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone). Lihat apps-script/PANDUAN.md.
 *
 * Aplikasi di GitHub Pages memanggil URL web app ini. Setiap permintaan
 * harus membawa PIN yang sama dengan isian PIN di sheet Pengaturan.
 */

var SHEET = { PENGATURAN: "Pengaturan", PEGAWAI: "Pegawai", POSYANDU: "Posyandu", KEGIATAN: "Kegiatan", PENUGASAN: "Penugasan" };

var KOLOM_PEGAWAI = ["nama", "nip", "pangkat", "jabatan"];
var KOLOM_POSYANDU = ["nama", "desa"];
var KOLOM_KEGIATAN = ["id", "dibuat", "diubah", "nomorST", "tglKeluar", "tglSPD", "maksud", "dasar", "kodeSPD",
  "tempatKeluar", "penggunaAnggaran", "tingkatBiaya", "alatAngkut", "tempatBerangkat", "skpd", "kodeRekening",
  "keterangan", "pejabatNama", "pejabatNip", "program", "kegiatanAnggaran", "subKegiatan", "uangTransport",
  "programVerifikasi", "pptkNama", "pptkNip", "verifKelengkapan", "verifPersiapan", "verifPptkUkm",
  "verifBendahara", "verifPjKeuangan"];
var KOLOM_PENUGASAN = ["kegiatanId", "nomor", "nama", "tujuan", "tanggal", "tglSPD"];

function doGet() {
  return json({ ok: true, pesan: "API SPD Moncongloe aktif. Buka aplikasi untuk memakainya." });
}

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ ok: false, pesan: "Permintaan tidak valid." });
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pengaturan = bacaPengaturan(ss);
  if (!pengaturan.PIN || String(req.pin || "") !== String(pengaturan.PIN)) {
    return json({ ok: false, kode: "PIN", pesan: "PIN salah. Periksa PIN di menu Pengaturan aplikasi." });
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (req.aksi) {
      case "muat": return json(muat(ss, pengaturan));
      case "ambil": return json(ambil(ss, req.id));
      case "simpanKegiatan": return json(simpanKegiatan(ss, req.kegiatan, req.penugasan || [], req.paksa));
      case "hapusKegiatan": return json(hapusKegiatan(ss, req.id));
      case "simpanPegawai": return json(simpanPegawai(ss, req.pegawai || []));
      case "simpanPosyandu": return json(simpanPosyandu(ss, req.posyandu || []));
      case "simpanPengaturan": return json(simpanPengaturan(ss, req.pengaturan || {}));
      default: return json({ ok: false, pesan: "Aksi tidak dikenal: " + req.aksi });
    }
  } catch (err) {
    return json({ ok: false, pesan: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- aksi
function muat(ss, pengaturan) {
  var tanpaPin = {};
  Object.keys(pengaturan).forEach(function (k) { if (k !== "PIN") tanpaPin[k] = pengaturan[k]; });
  var penugasan = bacaTabel(sheet(ss, SHEET.PENUGASAN), KOLOM_PENUGASAN);
  var ringkas = {};
  penugasan.forEach(function (p) {
    var r = ringkas[p.kegiatanId] = ringkas[p.kegiatanId] || { kunjungan: 0, pegawai: {} };
    r.kunjungan++;
    r.pegawai[p.nama.toLowerCase()] = 1;
  });
  var kegiatan = bacaTabel(sheet(ss, SHEET.KEGIATAN), KOLOM_KEGIATAN).map(function (k) {
    var r = ringkas[k.id] || { kunjungan: 0, pegawai: {} };
    return { id: k.id, nomorST: k.nomorST, tglKeluar: k.tglKeluar, maksud: k.maksud, diubah: k.diubah,
      jumlahKunjungan: r.kunjungan, jumlahPegawai: Object.keys(r.pegawai).length };
  });
  return { ok: true, pengaturan: tanpaPin, pegawai: bacaTabel(sheet(ss, SHEET.PEGAWAI), KOLOM_PEGAWAI),
    posyandu: bacaTabel(sheetPosyandu(ss), KOLOM_POSYANDU), kegiatan: kegiatan };
}

function ambil(ss, id) {
  var k = bacaTabel(sheet(ss, SHEET.KEGIATAN), KOLOM_KEGIATAN).filter(function (x) { return x.id === id; })[0];
  if (!k) throw new Error("Kegiatan tidak ditemukan (mungkin sudah dihapus).");
  var p = bacaTabel(sheet(ss, SHEET.PENUGASAN), KOLOM_PENUGASAN).filter(function (x) { return x.kegiatanId === id; });
  return { ok: true, kegiatan: k, penugasan: p };
}

function simpanKegiatan(ss, kegiatan, penugasan, paksa) {
  if (!kegiatan) throw new Error("Data kegiatan kosong.");
  var shK = sheet(ss, SHEET.KEGIATAN), semua = bacaTabel(shK, KOLOM_KEGIATAN);
  var sekarang = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var id = kegiatan.id;
  var lama = id ? semua.filter(function (x) { return x.id === id; })[0] : null;
  if (!lama) id = kegiatan.id || Utilities.getUuid().slice(0, 8);
  // kegiatan.diubah = versi yang dibuka petugas. Bila berbeda, sudah ada yang menyimpan lebih dulu.
  if (lama && !paksa && kegiatan.diubah && kegiatan.diubah !== lama.diubah) {
    return { ok: false, kode: "KONFLIK", diubah: lama.diubah,
      pesan: "Kegiatan ini sudah diubah orang lain pada " + lama.diubah + "." };
  }
  var baru = {};
  KOLOM_KEGIATAN.forEach(function (c) { baru[c] = kegiatan[c] == null ? "" : String(kegiatan[c]); });
  baru.id = id;
  baru.dibuat = lama ? lama.dibuat : sekarang;
  baru.diubah = sekarang;
  if (lama) semua = semua.map(function (x) { return x.id === id ? baru : x; });
  else semua.push(baru);
  tulisTabel(shK, KOLOM_KEGIATAN, semua);

  var shP = sheet(ss, SHEET.PENUGASAN);
  var lain = bacaTabel(shP, KOLOM_PENUGASAN).filter(function (x) { return x.kegiatanId !== id; });
  penugasan.forEach(function (p) {
    lain.push({ kegiatanId: id, nomor: String(p.nomor || ""), nama: String(p.nama || ""),
      tujuan: String(p.tujuan || ""), tanggal: String(p.tanggal || ""), tglSPD: String(p.tglSPD || "") });
  });
  tulisTabel(shP, KOLOM_PENUGASAN, lain);
  return { ok: true, id: id, diubah: sekarang };
}

function hapusKegiatan(ss, id) {
  var shK = sheet(ss, SHEET.KEGIATAN);
  tulisTabel(shK, KOLOM_KEGIATAN, bacaTabel(shK, KOLOM_KEGIATAN).filter(function (x) { return x.id !== id; }));
  var shP = sheet(ss, SHEET.PENUGASAN);
  tulisTabel(shP, KOLOM_PENUGASAN, bacaTabel(shP, KOLOM_PENUGASAN).filter(function (x) { return x.kegiatanId !== id; }));
  return { ok: true };
}

function simpanPegawai(ss, pegawai) {
  var bersih = pegawai.filter(function (p) { return p && String(p.nama || "").trim(); }).map(function (p) {
    return { nama: String(p.nama).trim(), nip: String(p.nip || "").replace(/\s+/g, ""),
      pangkat: String(p.pangkat || "").trim() || "-", jabatan: String(p.jabatan || "").trim() || "-" };
  });
  bersih.sort(function (a, b) { return a.nama.toLowerCase().localeCompare(b.nama.toLowerCase()); });
  tulisTabel(sheet(ss, SHEET.PEGAWAI), KOLOM_PEGAWAI, bersih);
  return { ok: true, pegawai: bersih };
}

function simpanPosyandu(ss, posyandu) {
  var bersih = posyandu.filter(function (p) { return p && String(p.nama || "").trim(); }).map(function (p) {
    return { nama: String(p.nama).trim(), desa: String(p.desa || "").trim() };
  });
  bersih.sort(function (a, b) { return a.nama.toLowerCase().localeCompare(b.nama.toLowerCase()); });
  tulisTabel(sheetPosyandu(ss), KOLOM_POSYANDU, bersih);
  return { ok: true, posyandu: bersih };
}

function simpanPengaturan(ss, nilai) {
  var sh = sheet(ss, SHEET.PENGATURAN);
  var data = sh.getDataRange().getDisplayValues();
  Object.keys(nilai).forEach(function (kunci) {
    if (kunci === "PIN") return;              // PIN hanya diubah langsung di sheet
    var ketemu = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === kunci) { teksSel(sh.getRange(i + 1, 2), nilai[kunci]); ketemu = true; break; }
    }
    if (!ketemu) {
      data.push([kunci]);
      sh.getRange(data.length, 1, 1, 2).setNumberFormat("@").setValues([[kunci, String(nilai[kunci])]]);
    }
  });
  return { ok: true };
}

// ---------------------------------------------------------------- util
// Format teks dulu supaya NIP panjang tidak diubah Google Sheet menjadi 1.98E+17.
function teksSel(rng, nilai) {
  rng.setNumberFormat("@").setValue(String(nilai == null ? "" : nilai));
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheet(ss, nama) {
  var sh = ss.getSheetByName(nama);
  if (!sh) throw new Error("Sheet '" + nama + "' tidak ditemukan. Pakai template Database SPD Moncongloe.");
  return sh;
}

// Sheet Posyandu dibuat otomatis bila Google Sheet lama belum memilikinya.
function sheetPosyandu(ss) {
  var sh = ss.getSheetByName(SHEET.POSYANDU);
  if (!sh) {
    sh = ss.insertSheet(SHEET.POSYANDU);
    sh.getRange(1, 1, 1, 2).setValues([KOLOM_POSYANDU]);
  }
  return sh;
}

function bacaPengaturan(ss) {
  var hasil = {};
  sheet(ss, SHEET.PENGATURAN).getDataRange().getDisplayValues().slice(1).forEach(function (r) {
    if (r[0]) hasil[String(r[0]).trim()] = String(r[1]);
  });
  return hasil;
}

// Baris 1 = nama kolom. Semua nilai dibaca sebagai teks apa adanya.
function bacaTabel(sh, kolom) {
  var data = sh.getDataRange().getDisplayValues();
  if (data.length < 2) return [];
  var kepala = data[0].map(function (h) { return String(h).trim(); });
  var idx = kolom.map(function (c) { return kepala.indexOf(c); });
  return data.slice(1).filter(function (r) { return r.join("") !== ""; }).map(function (r) {
    var o = {};
    kolom.forEach(function (c, i) { o[c] = idx[i] >= 0 ? String(r[idx[i]]) : ""; });
    return o;
  });
}

function tulisTabel(sh, kolom, baris) {
  var nilai = [kolom].concat(baris.map(function (o) { return kolom.map(function (c) { return o[c] == null ? "" : String(o[c]); }); }));
  var lebar = Math.max(sh.getMaxColumns(), kolom.length);
  sh.getRange(1, 1, sh.getMaxRows(), lebar).clearContent();
  if (sh.getMaxRows() < nilai.length) sh.insertRowsAfter(sh.getMaxRows(), nilai.length - sh.getMaxRows());
  var rng = sh.getRange(1, 1, nilai.length, kolom.length);
  rng.setNumberFormat("@");                   // simpan sebagai teks: NIP & tanggal tidak berubah
  rng.setValues(nilai);
}
