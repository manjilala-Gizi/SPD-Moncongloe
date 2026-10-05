/*
 * SPD Core — membaca Template Input SPD, memeriksa isinya, dan
 * menghasilkan file SPD (satu sheet per SPD) dari cetakan master_spd.xlsx.
 *
 * Berjalan sepenuhnya di browser (juga bisa diuji di Node). Tidak ada data
 * yang dikirim ke server mana pun.
 *
 * Ketergantungan disuntikkan lewat parameter: XLSX (SheetJS, untuk membaca),
 * JSZip (untuk menulis), DOMParser & XMLSerializer.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.SPDCore = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var MAKS_KUNJUNGAN = 6;               // slot II–VII di halaman belakang SPD
  var BARIS_SLOT = [10, 15, 20, 25, 30, 35];
  var BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
    "Agustus", "September", "Oktober", "November", "Desember"];

  // ------------------------------------------------------------ util teks
  function teks(v) {
    if (v === null || v === undefined) return "";
    return String(v).replace(/\s+/g, " ").trim();
  }
  function kunciNama(v) { return teks(v).toLowerCase(); }
  function pad2(n) { return (n < 10 ? "0" : "") + n; }

  var SATUAN = ["", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh",
    "Delapan", "Sembilan", "Sepuluh", "Sebelas"];
  function terbilang(n) {
    if (n < 12) return SATUAN[n];
    if (n < 20) return SATUAN[n - 10] + " Belas";
    if (n < 100) return SATUAN[Math.floor(n / 10)] + " Puluh" + (n % 10 ? " " + SATUAN[n % 10] : "");
    return String(n);
  }

  // ------------------------------------------------------------ tanggal
  // Tanggal disimpan sebagai {y, m (1-12), d}. Serial Excel dihitung dengan
  // UTC supaya tidak bergeser karena zona waktu perangkat.
  function dariSerial(serial) {
    var ms = Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000;
    var dt = new Date(ms);
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
  }
  function bacaTanggal(v) {
    if (v === null || v === undefined || v === "") return { tgl: null };
    if (typeof v === "number" && isFinite(v) && v > 20000 && v < 80000) return { tgl: dariSerial(v) };
    if (v instanceof Date && !isNaN(v)) return { tgl: { y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate() } };
    var m = /^(\d{1,2})[-\/. ](\d{1,2})[-\/. ](\d{4})$/.exec(teks(v));
    if (m) {
      var t = { y: +m[3], m: +m[2], d: +m[1] };
      var cek = new Date(Date.UTC(t.y, t.m - 1, t.d));
      if (cek.getUTCMonth() + 1 === t.m && cek.getUTCDate() === t.d) return { tgl: t, sebagaiTeks: true };
    }
    return { tgl: null, salah: true };
  }
  function kunciTgl(t) { return t.y + "-" + pad2(t.m) + "-" + pad2(t.d); }
  function tglPanjang(t) { return pad2(t.d) + " " + BULAN[t.m - 1] + " " + t.y; }

  // "06, 13, 15 Agustus 2026" / "28, 30 Juli, 02 Agustus 2026" /
  // "30 Desember 2026, 02 Januari 2027"
  function daftarTanggal(list) {
    var grup = [];
    list.forEach(function (t) {
      var g = grup[grup.length - 1];
      if (g && g.y === t.y && g.m === t.m) g.d.push(t.d);
      else grup.push({ y: t.y, m: t.m, d: [t.d] });
    });
    var satuTahun = grup.every(function (g) { return g.y === grup[0].y; });
    var bagian = grup.map(function (g, i) {
      var s = g.d.map(pad2).join(", ") + " " + BULAN[g.m - 1];
      if (!satuTahun || i === grup.length - 1) s += " " + g.y;
      return s;
    });
    return bagian.join(", ");
  }

  // ------------------------------------------------------------ membaca
  var LABEL_KEGIATAN = [
    ["nomor surat tugas", "nomorST"],
    ["tanggal dikeluarkan", "tglKeluar"],
    ["tempat dikeluarkan", "tempatKeluar"],
    ["pengguna anggaran", "penggunaAnggaran"],
    ["maksud perjalanan", "maksud"],
    ["tingkat biaya", "tingkatBiaya"],
    ["alat angkut", "alatAngkut"],
    ["tempat berangkat", "tempatBerangkat"],
    ["skpd", "skpd"],
    ["kode rekening", "kodeRekening"],
    ["keterangan lain", "keterangan"],
    ["nama pejabat", "pejabatNama"],
    ["nip pejabat", "pejabatNip"]
  ];
  var WAJIB_KEGIATAN = {
    tglKeluar: "Tanggal dikeluarkan (SPD)", tempatKeluar: "Tempat dikeluarkan",
    penggunaAnggaran: "Pengguna Anggaran/Kuasa Pengguna Anggaran", maksud: "Maksud Perjalanan Dinas",
    tingkatBiaya: "Tingkat Biaya Perjalanan Dinas", alatAngkut: "Alat Angkut", tempatBerangkat: "Tempat Berangkat",
    skpd: "SKPD", kodeRekening: "Kode Rekening", pejabatNama: "Nama Pejabat Penandatangan",
    pejabatNip: "NIP Pejabat Penandatangan"
  };

  function cariSheet(wb, nama) {
    var n = wb.SheetNames.filter(function (s) { return s.trim().toLowerCase() === nama.toLowerCase(); })[0];
    return n ? wb.Sheets[n] : null;
  }
  function baris(XLSX, ws) {
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  }

  function bacaPegawai(XLSX, ws, catatan) {
    var hasil = [];
    if (!ws) return hasil;
    baris(XLSX, ws).forEach(function (r, i) {
      if (i === 0) return;
      var nama = teks(r[0]);
      if (!nama) return;
      var selNip = ws[XLSX.utils.encode_cell({ r: i, c: 1 })];
      if (catatan && selNip && selNip.t === "n" && String(Math.round(selNip.v)).length > 15) {
        catatan.push("NIP/NIK " + nama + " tersimpan sebagai angka sehingga digit terakhirnya bisa berubah. " +
          "Ketik ulang di sheet Pegawai sebagai teks (awali dengan tanda petik ').");
      }
      var nip = selNip && selNip.t === "n" ? String(Math.round(selNip.v)) : teks(r[1]);
      hasil.push({ nama: nama, nip: nip, pangkat: teks(r[2]), jabatan: teks(r[3]), baris: i + 1 });
    });
    return hasil;
  }

  /** Membaca file template (ArrayBuffer) menjadi data mentah. */
  function bacaTemplate(XLSX, buf) {
    var wb = XLSX.read(buf, { type: "array" });
    var catatan = [];
    var wsK = cariSheet(wb, "Kegiatan"), wsP = cariSheet(wb, "Penugasan"), wsG = cariSheet(wb, "Pegawai");
    var hilang = [];
    if (!wsK) hilang.push("Kegiatan");
    if (!wsP) hilang.push("Penugasan");
    if (!wsG) hilang.push("Pegawai");
    if (hilang.length) {
      throw new Error("File ini bukan Template Input SPD: sheet " + hilang.join(", ") +
        " tidak ditemukan. Unduh template dari aplikasi lalu isi ulang.");
    }

    var kegiatan = {}, tglMentah = null;
    baris(XLSX, wsK).forEach(function (r) {
      var label = teks(r[0]).toLowerCase();
      for (var i = 0; i < LABEL_KEGIATAN.length; i++) {
        if (label.indexOf(LABEL_KEGIATAN[i][0]) === 0 && !(LABEL_KEGIATAN[i][1] in kegiatan)) {
          if (LABEL_KEGIATAN[i][1] === "tglKeluar") tglMentah = r[1];
          kegiatan[LABEL_KEGIATAN[i][1]] = teks(r[1]);
          break;
        }
      }
    });
    kegiatan.tglKeluarRaw = tglMentah;

    var penugasan = [];
    baris(XLSX, wsP).forEach(function (r, i) {
      if (i === 0) return;
      var isi = [r[0], r[1], r[2], r[3]];
      if (isi.every(function (v) { return teks(v) === ""; })) return;
      penugasan.push({ baris: i + 1, nomor: teks(r[0]), nama: teks(r[1]), tujuan: teks(r[2]), tanggalRaw: r[3] });
    });

    return { kegiatan: kegiatan, pegawai: bacaPegawai(XLSX, wsG, catatan), penugasan: penugasan, catatan: catatan };
  }

  /** Daftar pegawai dari template resmi (repo) untuk pembanding. */
  function bacaPegawaiResmi(XLSX, buf) {
    var wb = XLSX.read(buf, { type: "array" });
    return bacaPegawai(XLSX, cariSheet(wb, "Pegawai"), null);
  }

  // ------------------------------------------------------------ pemeriksaan
  function nomorUrut(nomor) {
    var m = /^\s*(\d+)/.exec(nomor);
    return m ? +m[1] : Infinity;
  }

  /**
   * Memeriksa data dan menyusun rencana SPD.
   * Hasil: { galat[], peringatan[], spd[], perluNomor[] }
   * galat = menghentikan proses; peringatan = boleh lanjut.
   */
  function periksa(data, pegawaiResmi) {
    var galat = [], peringatan = data.catatan.slice();
    var k = data.kegiatan;

    Object.keys(WAJIB_KEGIATAN).forEach(function (kunci) {
      if (!k[kunci]) galat.push("Sheet Kegiatan: '" + WAJIB_KEGIATAN[kunci] + "' belum diisi.");
    });
    if (!k.nomorST) peringatan.push("Sheet Kegiatan: Nomor Surat Tugas kosong (tidak dicetak di SPD, hanya sebagai rujukan).");
    if (!k.keterangan) k.keterangan = "-";

    var tk = bacaTanggal(k.tglKeluarRaw);
    if (k.tglKeluar && !tk.tgl) galat.push("Sheet Kegiatan: Tanggal dikeluarkan bukan tanggal yang valid.");
    k.tgl = tk.tgl;

    var petaPegawai = {};
    data.pegawai.forEach(function (p) {
      var key = kunciNama(p.nama);
      if (petaPegawai[key]) peringatan.push("Sheet Pegawai: nama '" + p.nama + "' tercantum lebih dari sekali; dipakai yang pertama.");
      else petaPegawai[key] = p;
    });
    if (!data.pegawai.length) galat.push("Sheet Pegawai kosong.");
    if (!data.penugasan.length) galat.push("Sheet Penugasan belum berisi kunjungan.");

    // baris-baris kunjungan
    var kunjungan = [];
    data.penugasan.forEach(function (r) {
      var lok = "Penugasan baris " + r.baris + ": ";
      var masalah = [];
      if (!r.nama) masalah.push("nama pegawai kosong");
      var peg = r.nama ? petaPegawai[kunciNama(r.nama)] : null;
      if (r.nama && !peg) masalah.push("nama '" + r.nama + "' tidak ada di sheet Pegawai");
      if (!r.nomor) masalah.push("Nomor SPD kosong");
      if (!r.tujuan) masalah.push("Tempat Tujuan kosong");
      var t = bacaTanggal(r.tanggalRaw);
      if (!t.tgl) masalah.push(t.salah ? "Tanggal tidak dikenali ('" + teks(r.tanggalRaw) + "')" : "Tanggal kosong");
      else if (t.sebagaiTeks) peringatan.push(lok + "tanggal diketik sebagai teks, dibaca sebagai " + tglPanjang(t.tgl) + ".");
      if (masalah.length) { galat.push(lok + masalah.join("; ") + "."); return; }
      if (k.tgl && kunciTgl(t.tgl) < kunciTgl(k.tgl)) {
        peringatan.push(lok + "kunjungan " + tglPanjang(t.tgl) + " lebih awal dari tanggal SPD dikeluarkan (" + tglPanjang(k.tgl) + ").");
      }
      kunjungan.push({ baris: r.baris, nomor: r.nomor, pegawai: peg, tujuan: r.tujuan, tgl: t.tgl });
    });

    // satu nomor = satu pegawai, satu pegawai = satu nomor
    var perNomor = {}, nomorPerPegawai = {};
    kunjungan.forEach(function (v) {
      (perNomor[v.nomor] = perNomor[v.nomor] || []).push(v);
      var kp = kunciNama(v.pegawai.nama);
      (nomorPerPegawai[kp] = nomorPerPegawai[kp] || {})[v.nomor] = true;
    });
    Object.keys(perNomor).forEach(function (no) {
      var nama = {};
      perNomor[no].forEach(function (v) { nama[v.pegawai.nama] = true; });
      var n = Object.keys(nama);
      if (n.length > 1) galat.push("Nomor SPD " + no + " dipakai oleh lebih dari satu pegawai (" + n.join("; ") + ").");
    });
    Object.keys(nomorPerPegawai).forEach(function (kp) {
      var n = Object.keys(nomorPerPegawai[kp]);
      if (n.length > 1) {
        galat.push(petaPegawai[kp].nama + " memakai " + n.length + " Nomor SPD berbeda (" + n.join(", ") +
          "). Pakai satu nomor saja; bila kunjungannya lebih dari 6, aplikasi akan memecahnya.");
      }
    });

    // rencana SPD
    var spd = [], perluNomor = [];
    Object.keys(perNomor).forEach(function (no) {
      var v = perNomor[no].slice().sort(function (a, b) {
        var x = kunciTgl(a.tgl), y = kunciTgl(b.tgl);
        return x < y ? -1 : x > y ? 1 : a.baris - b.baris;
      });
      var peg = v[0].pegawai;
      var seen = {};
      v.forEach(function (x) {
        var kk = kunciTgl(x.tgl);
        if (seen[kk]) {
          peringatan.push(peg.nama + ": dua kunjungan pada tanggal yang sama (" + tglPanjang(x.tgl) +
            ", baris " + seen[kk] + " dan " + x.baris + ").");
        } else seen[kk] = x.baris;
      });
      if (!peg.nip) peringatan.push("Sheet Pegawai: NIP/NIK " + peg.nama + " kosong.");
      if (!peg.jabatan) peringatan.push("Sheet Pegawai: Jabatan " + peg.nama + " kosong.");

      var jumlahBagian = Math.ceil(v.length / MAKS_KUNJUNGAN);
      for (var b = 0; b < jumlahBagian; b++) {
        var item = {
          kunci: no + "#" + (b + 1), nomorAsal: no, bagian: b + 1, jumlahBagian: jumlahBagian,
          nomor: b === 0 ? no : null, pegawai: peg,
          kunjungan: v.slice(b * MAKS_KUNJUNGAN, (b + 1) * MAKS_KUNJUNGAN)
        };
        spd.push(item);
        if (b > 0) perluNomor.push(item);
      }
      if (jumlahBagian > 1) {
        peringatan.push(peg.nama + " punya " + v.length + " kunjungan, dipecah menjadi " + jumlahBagian +
          " SPD (maks. 6 kunjungan per SPD). Isi Nomor SPD tambahan di bawah.");
      }
    });
    spd.sort(function (a, b) {
      var x = nomorUrut(a.nomorAsal), y = nomorUrut(b.nomorAsal);
      if (x !== y) return x - y;
      if (a.nomorAsal !== b.nomorAsal) return a.nomorAsal < b.nomorAsal ? -1 : 1;
      return a.bagian - b.bagian;
    });

    // pembanding daftar pegawai resmi (repo)
    if (pegawaiResmi && pegawaiResmi.length) {
      var beda = bandingkanPegawai(data.pegawai, pegawaiResmi);
      if (beda.length) {
        peringatan.push("Daftar pegawai di file ini berbeda dengan template terbaru di aplikasi (" + beda.join("; ") +
          "). SPD tetap dibuat dari data di file Anda. Unduh template terbaru agar data pegawai selalu mutakhir.");
      }
    }

    return { galat: galat, peringatan: peringatan, spd: spd, perluNomor: perluNomor, kegiatan: k };
  }

  function bandingkanPegawai(milikFile, resmi) {
    var a = {}, b = {}, beda = [];
    milikFile.forEach(function (p) { a[kunciNama(p.nama)] = p; });
    resmi.forEach(function (p) { b[kunciNama(p.nama)] = p; });
    var tambah = resmi.filter(function (p) { return !a[kunciNama(p.nama)]; }).map(function (p) { return p.nama; });
    var hapus = milikFile.filter(function (p) { return !b[kunciNama(p.nama)]; }).map(function (p) { return p.nama; });
    var ubah = resmi.filter(function (p) {
      var q = a[kunciNama(p.nama)];
      return q && (q.nip !== p.nip || q.pangkat !== p.pangkat || q.jabatan !== p.jabatan);
    }).map(function (p) { return p.nama; });
    if (tambah.length) beda.push("belum ada: " + tambah.join(", "));
    if (hapus.length) beda.push("tidak ada di versi terbaru: " + hapus.join(", "));
    if (ubah.length) beda.push("data berubah: " + ubah.join(", "));
    return beda;
  }

  // ------------------------------------------------------------ isi sel
  function isiSel(k, item) {
    var p = item.kunjungan, peg = item.pegawai, c = {};
    var tujuan = [];
    p.forEach(function (v) { if (tujuan.indexOf(v.tujuan) < 0) tujuan.push(v.tujuan); });
    var tgl = daftarTanggal(p.map(function (v) { return v.tgl; }));
    var nipPejabat = "NIP. " + k.pejabatNip;

    // ---- halaman depan
    c.G10 = ": " + item.nomor;
    c.D14 = k.penggunaAnggaran;
    c.D15 = peg.nama;
    c.D16 = peg.nip;
    c.E17 = peg.pangkat || "-";
    c.E18 = peg.jabatan || "-";
    c.E19 = k.tingkatBiaya;
    c.D20 = k.maksud;
    c.D22 = k.alatAngkut;
    c.E23 = k.tempatBerangkat;
    c.E24 = tujuan.join(", ");
    c.E26 = p.length + " (" + terbilang(p.length) + ") Hari";
    c.E27 = tgl;
    c.E28 = tgl;
    c.D36 = k.skpd;
    c.D37 = k.kodeRekening;
    c.D38 = k.keterangan;
    c.G39 = ": " + k.tempatKeluar;
    c.G40 = ": " + tglPanjang(k.tgl);
    c.F42 = k.penggunaAnggaran;
    c.F47 = k.pejabatNama;
    c.F48 = nipPejabat;

    // ---- halaman belakang
    c.P1 = k.tempatBerangkat;
    c.P3 = p[0].tujuan;
    c.P4 = tglPanjang(p[0].tgl);
    c.N5 = k.penggunaAnggaran;
    c.N8 = k.pejabatNama;
    c.N9 = nipPejabat;
    BARIS_SLOT.forEach(function (r, i) {
      var v = p[i];
      c["L" + r] = v ? v.tujuan : "";
      c["L" + (r + 1)] = v ? tglPanjang(v.tgl) : "";
      c["P" + r] = v ? v.tujuan : "";
      c["P" + (r + 1)] = v ? k.tempatBerangkat : "";
      c["P" + (r + 2)] = v ? tglPanjang(v.tgl) : "";
    });
    c.L40 = k.tempatBerangkat;
    c.L41 = tglPanjang(p[p.length - 1].tgl);
    c.J42 = k.penggunaAnggaran;
    c.J43 = k.pejabatNama;
    c.J46 = nipPejabat;
    return c;
  }

  // ------------------------------------------------------------ XML sheet
  var NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  var XMLNS = "http://www.w3.org/XML/1998/namespace";

  function kolom(ref) {
    var s = /^[A-Z]+/.exec(ref)[0], n = 0;
    for (var i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64;
    return n;
  }
  function anakElemen(el, nama) {
    var out = [];
    for (var n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && n.localName === nama) out.push(n);
    return out;
  }

  function tulisSel(doc, sheetData, ref, nilai) {
    var r = +/\d+$/.exec(ref)[0], col = kolom(ref);
    var rowEl = null, sebelum = null;
    var rows = anakElemen(sheetData, "row");
    for (var i = 0; i < rows.length; i++) {
      var rr = +rows[i].getAttribute("r");
      if (rr === r) { rowEl = rows[i]; break; }
      if (rr > r) { sebelum = rows[i]; break; }
    }
    if (!rowEl) {
      rowEl = doc.createElementNS(NS, "row");
      rowEl.setAttribute("r", String(r));
      sheetData.insertBefore(rowEl, sebelum);
    }
    var sel = null, selSebelum = null, cells = anakElemen(rowEl, "c");
    for (var j = 0; j < cells.length; j++) {
      var cc = kolom(cells[j].getAttribute("r"));
      if (cc === col) { sel = cells[j]; break; }
      if (cc > col) { selSebelum = cells[j]; break; }
    }
    if (!sel) {
      sel = doc.createElementNS(NS, "c");
      sel.setAttribute("r", ref);
      rowEl.insertBefore(sel, selSebelum);
    }
    while (sel.firstChild) sel.removeChild(sel.firstChild);
    sel.removeAttribute("t");
    if (nilai === "" || nilai === null || nilai === undefined) return;
    sel.setAttribute("t", "inlineStr");
    var is = doc.createElementNS(NS, "is"), t = doc.createElementNS(NS, "t");
    t.setAttributeNS(XMLNS, "xml:space", "preserve");
    t.appendChild(doc.createTextNode(String(nilai)));
    is.appendChild(t);
    sel.appendChild(is);
  }

  // Sel tujuan (E24:G25) muat ±3 baris teks. Bila daftar tujuan lebih
  // panjang, baris 25 ditinggikan supaya teks tidak terpotong di Excel.
  var MUAT_PER_BARIS = 70, TINGGI_BARIS_TEKS = 13;
  function tinggiTambahan(teksTujuan) {
    var baris = Math.ceil(teksTujuan.length / MUAT_PER_BARIS);
    return baris > 3 ? (baris - 3) * TINGGI_BARIS_TEKS : 0;
  }

  function isiSheetXml(env, xml, sel, aktif) {
    var doc = new env.DOMParser().parseFromString(xml, "application/xml");
    var sheetData = doc.getElementsByTagName("sheetData")[0];
    Object.keys(sel).forEach(function (ref) { tulisSel(doc, sheetData, ref, sel[ref]); });
    var tambah = tinggiTambahan(sel.E24 || "");
    if (tambah) {
      anakElemen(sheetData, "row").forEach(function (r) {
        if (r.getAttribute("r") === "25") {
          r.setAttribute("ht", String(parseFloat(r.getAttribute("ht") || "15") + tambah));
          r.setAttribute("customHeight", "1");
        }
      });
    }
    var sv = doc.getElementsByTagName("sheetView")[0];
    if (sv && !aktif) sv.removeAttribute("tabSelected");
    var out = new env.XMLSerializer().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, "");
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + out;
  }

  // ------------------------------------------------------------ nama sheet & file
  function namaDepan(nama) {
    var s = teks(nama).split(",")[0].trim().split(/\s+/)[0] || "Pegawai";
    return s.replace(/[\\\/\?\*\[\]:']/g, "");
  }
  function namaSheet(item, dipakai) {
    var t = item.kunjungan[0].tgl;
    var dasar = (namaDepan(item.pegawai.nama) + "_" + pad2(t.d) + "-" + pad2(t.m)).slice(0, 31);
    var nama = dasar, i = 2;
    while (dipakai[nama.toLowerCase()]) {
      var akhiran = " (" + i++ + ")";
      nama = dasar.slice(0, 31 - akhiran.length) + akhiran;
    }
    dipakai[nama.toLowerCase()] = true;
    return nama;
  }
  function namaFile(k) {
    var s = "SPD " + k.maksud + (k.tgl ? " " + BULAN[k.tgl.m - 1] + " " + k.tgl.y : "");
    return s.replace(/[\\\/:\*\?"<>\|]/g, "-").replace(/\s+/g, " ").trim() + ".xlsx";
  }
  function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  /**
   * Menghasilkan file SPD.
   * env: { JSZip, DOMParser, XMLSerializer }
   * nomorTambahan: { "<kunci>": "nomor SPD" } untuk bagian ke-2 dst.
   * Hasil: { data: Uint8Array, namaFile, sheet: [{nama, nomor, pegawai, jumlah}] }
   */
  async function buatSPD(env, masterBuf, hasilPeriksa, nomorTambahan) {
    if (hasilPeriksa.galat.length) throw new Error("Masih ada kesalahan pada data.");
    var k = hasilPeriksa.kegiatan;
    var daftar = hasilPeriksa.spd.map(function (it) {
      var o = Object.assign({}, it);
      if (!o.nomor) o.nomor = teks((nomorTambahan || {})[o.kunci]);
      if (!o.nomor) throw new Error("Nomor SPD tambahan untuk " + o.pegawai.nama + " (bagian " + o.bagian + ") belum diisi.");
      return o;
    });
    var semuaNomor = {};
    daftar.forEach(function (o) {
      if (semuaNomor[o.nomor]) throw new Error("Nomor SPD " + o.nomor + " dipakai dua kali.");
      semuaNomor[o.nomor] = true;
    });

    var zip = await env.JSZip.loadAsync(masterBuf);
    var tplSheet = await zip.file("xl/worksheets/sheet1.xml").async("string");
    var tplDrawing = await zip.file("xl/drawings/drawing1.xml").async("string");
    var tplDrawingRels = await zip.file("xl/drawings/_rels/drawing1.xml.rels").async("string");
    var ct = await zip.file("[Content_Types].xml").async("string");
    var wb = await zip.file("xl/workbook.xml").async("string");
    var wbRels = await zip.file("xl/_rels/workbook.xml.rels").async("string");

    ["xl/worksheets/sheet1.xml", "xl/worksheets/_rels/sheet1.xml.rels",
      "xl/drawings/drawing1.xml", "xl/drawings/_rels/drawing1.xml.rels"].forEach(function (p) { zip.remove(p); });
    ct = ct.replace(/<Override PartName="\/xl\/(worksheets\/sheet|drawings\/drawing)\d+\.xml"[^>]*\/>/g, "");
    wbRels = wbRels.replace(/<Relationship [^>]*relationships\/worksheet"[^>]*\/>/g, "");

    var dipakai = {}, sheetsXml = "", relsXml = "", ctXml = "", ringkasan = [];
    daftar.forEach(function (o, i) {
      var n = i + 1, nama = namaSheet(o, dipakai);
      zip.file("xl/worksheets/sheet" + n + ".xml", isiSheetXml(env, tplSheet, isiSel(k, o), i === 0));
      zip.file("xl/worksheets/_rels/sheet" + n + ".xml.rels",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing' + n + '.xml"/></Relationships>');
      zip.file("xl/drawings/drawing" + n + ".xml", tplDrawing);
      zip.file("xl/drawings/_rels/drawing" + n + ".xml.rels", tplDrawingRels);
      sheetsXml += '<sheet name="' + esc(nama) + '" sheetId="' + n + '" r:id="rIdS' + n + '"/>';
      relsXml += '<Relationship Id="rIdS' + n + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + n + '.xml"/>';
      ctXml += '<Override PartName="/xl/worksheets/sheet' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/drawings/drawing' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>';
      ringkasan.push({ nama: nama, nomor: o.nomor, pegawai: o.pegawai.nama, jumlah: o.kunjungan.length });
    });

    zip.file("xl/workbook.xml", wb.replace(/<sheets>.*<\/sheets>/, "<sheets>" + sheetsXml + "</sheets>"));
    zip.file("xl/_rels/workbook.xml.rels", wbRels.replace("</Relationships>", relsXml + "</Relationships>"));
    zip.file("[Content_Types].xml", ct.replace("</Types>", ctXml + "</Types>"));

    var data = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    return { data: data, namaFile: namaFile(k), sheet: ringkasan };
  }

  /** Nama sheet yang akan dipakai, untuk ditampilkan sebelum file dibuat. */
  function pratinjauNamaSheet(spd) {
    var dipakai = {};
    return spd.map(function (o) { return namaSheet(o, dipakai); });
  }

  return {
    pratinjauNamaSheet: pratinjauNamaSheet,
    daftarTanggal: daftarTanggal,
    bacaTemplate: bacaTemplate,
    bacaPegawaiResmi: bacaPegawaiResmi,
    periksa: periksa,
    buatSPD: buatSPD,
    _uji: { daftarTanggal: daftarTanggal, terbilang: terbilang, tglPanjang: tglPanjang, namaDepan: namaDepan, bacaTanggal: bacaTanggal }
  };
});
