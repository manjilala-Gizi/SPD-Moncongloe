/* Antarmuka SPD Moncongloe: dashboard, form kegiatan, data pegawai, pengaturan. */
(function () {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };
  var env = { JSZip: window.JSZip, DOMParser: window.DOMParser, XMLSerializer: window.XMLSerializer };
  var BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

  // Nilai bawaan bila Google Sheet belum diatur (sama dengan sheet Pengaturan).
  var BAWAAN = {
    tempatKeluar: "Moncongloe", penggunaAnggaran: "Kepala UPTD Puskesmas Moncongloe", tingkatBiaya: "Biasa",
    alatAngkut: "Kendaraan Roda Dua", tempatBerangkat: "UPTD Puskesmas Moncongloe", skpd: "UPTD Puskesmas Moncongloe",
    kodeRekening: "1 02 02 2.02 0015", keterangan: "-", pejabatNama: "Muliati, SKM, M.Kes", pejabatNip: "197811302015082001",
    kodeSPD: "SPD/PKM-ML",
    dasar: "Berdasarkan Rencana Pelaksanaan Kegiatan Tahun 2026 Bersumber Dana Biaya Operasional Kesehatan (BOK)",
    program: "1.02.02. - PROGRAM PEMENUHAN UPAYA KESEHATAN PERORANGAN DAN UPAYA KESEHATAN MASYARAKAT",
    kegiatanAnggaran: "1.02.02.2.02 - Penyediaan Layanan Kesehatan untuk UKM dan UKP Rujukan Tingkat Daerah Kabupaten/Kota",
    subKegiatan: "1 02 02 2.02 0015 Pengelolaan Pelayanan Kesehatan Gizi Masyarakat",
    uangTransport: "100000", programVerifikasi: "GIZI", pptkNama: "Roslin Salima, S.KM, M.Kes", pptkNip: "198203292009032006",
    verifKelengkapan: "Sulhiyah Mahmud, SKM, M.Kes", verifPersiapan: "Sitti Nurhayati, A.Md.RMIK",
    verifPptkUkm: "Ns. Ridwan Rahman, S.Kep.", verifBendahara: "Kameliah, SKM", verifPjKeuangan: "Roslin Salima, S.KM, M.Kes"
  };

  // Isian form kegiatan, dikelompokkan per bagian. [kunci, label, jenis]
  var BAGIAN_ST = [
    ["nomorST", "Nomor Surat Tugas", "text"], ["tglKeluar", "Tanggal Surat Tugas", "date"],
    ["tglSPD", "Tanggal SPD (kosongkan bila sama dengan ST)", "date"],
    ["maksud", "Nama kegiatan / Maksud perjalanan dinas", "text"], ["kodeSPD", "Kode Nomor SPD", "text"],
    ["dasar", "Dasar Surat Tugas", "textarea"]
  ];
  var BAGIAN_SPD = [
    ["tempatKeluar", "Tempat dikeluarkan", "text"], ["penggunaAnggaran", "Pengguna Anggaran", "text"],
    ["tingkatBiaya", "Tingkat biaya", "text"], ["alatAngkut", "Alat angkut", "text"],
    ["tempatBerangkat", "Tempat berangkat", "text"], ["skpd", "SKPD", "text"],
    ["kodeRekening", "Kode rekening", "text"], ["keterangan", "Keterangan lain-lain", "text"]
  ];
  var BAGIAN_TTD = [
    ["pejabatNama", "Nama pejabat penanda tangan", "text"], ["pejabatNip", "NIP pejabat", "text"],
    ["pptkNama", "Nama PPTK (Lembar Penerimaan)", "text"], ["pptkNip", "NIP PPTK", "text"]
  ];
  var BAGIAN_ANGGARAN = [
    ["program", "Program", "text"], ["kegiatanAnggaran", "Kegiatan (anggaran)", "text"],
    ["subKegiatan", "Sub Kegiatan", "text"], ["uangTransport", "Uang transport per kunjungan (Rp)", "number"],
    ["programVerifikasi", "Program (Lembar Verifikasi)", "text"]
  ];
  var BAGIAN_VERIF = [
    ["verifKelengkapan", "Verifikator Kelengkapan Dokumen", "text"], ["verifPersiapan", "Verifikator Persiapan Kegiatan", "text"],
    ["verifPptkUkm", "Verifikator PPTK UKM", "text"], ["verifBendahara", "Verifikator Bendahara", "text"],
    ["verifPjKeuangan", "Verifikator PJ Keuangan", "text"]
  ];
  var SEMUA_BAGIAN = BAGIAN_ST.concat(BAGIAN_SPD, BAGIAN_TTD, BAGIAN_ANGGARAN, BAGIAN_VERIF);
  var KUNCI_PENGATURAN = SEMUA_BAGIAN.map(function (x) { return x[0]; }).filter(function (k) {
    return ["nomorST", "tglKeluar", "tglSPD", "maksud"].indexOf(k) < 0;
  });

  var XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  var KELUARAN = {
    spd: { label: "SPD (Excel)", master: "assets/master_spd.xlsx", buat: "buatSPD", tipe: XLSX_MIME },
    st: { label: "Surat Tugas (Word)", master: "assets/master_st.docx", buat: "buatST",
      tipe: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
    lpj: { label: "Lembar Penerimaan (Excel)", master: "assets/master_penerimaan.xlsx", buat: "buatPenerimaan", tipe: XLSX_MIME },
    ver: { label: "Lembar Verifikasi (Excel)", master: "assets/master_verifikasi.xlsx", buat: "buatVerifikasi", tipe: XLSX_MIME }
  };

  var keadaan = {
    pengaturan: Object.assign({}, BAWAAN), pegawai: [], posyandu: [], kegiatan: [],
    status: "memuat", pesanStatus: "", form: null, berubah: false
  };

  // ---------------------------------------------------------------- util
  function esc(s) {
    return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  function toast(teks, gagal) {
    var t = $("toast");
    t.textContent = teks;
    t.className = "toast" + (gagal ? " gagal" : "");
    t.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { t.hidden = true; }, gagal ? 6000 : 3500);
  }
  function ambilFile(url) {
    return fetch(url, { cache: "no-cache" }).then(function (r) {
      if (!r.ok) throw new Error("Gagal memuat " + url + " (" + r.status + ")");
      return r.arrayBuffer();
    });
  }
  function unduh(data, nama, tipe) {
    var a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([data], { type: tipe }));
    a.download = nama;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
  }
  function pad2(n) { return (n < 10 ? "0" : "") + n; }
  function tglTampil(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || "");
    return m ? m[3] + "-" + m[2] + "-" + m[1] : (iso || "—");
  }
  function tglPanjang(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || "");
    return m ? +m[3] + " " + BULAN[+m[2] - 1] + " " + m[1] : "—";
  }
  function isoDari(raw) {
    var t = SPDCore._uji.bacaTanggal(raw).tgl;
    return t ? t.y + "-" + pad2(t.m) + "-" + pad2(t.d) : "";
  }
  function hariIni() { var d = new Date(); return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate()); }
  function kunciNama(s) { return String(s || "").toLowerCase().replace(/[^a-z0-9]/g, ""); }
  function petaPegawai() {
    var p = {};
    keadaan.pegawai.forEach(function (x) { p[kunciNama(x.nama)] = x; });
    return p;
  }
  // "tokka", "Tokka", atau "tokka desa bonto marannu" -> "Tokka Desa Bonto Marannu".
  function rapat(s) { return String(s || "").replace(/\s+/g, " ").trim().toLowerCase(); }
  function tujuanLengkap(p) { return p.nama + " Desa " + p.desa; }
  function cariPosyandu(teks) {
    var t = rapat(teks);
    if (!t) return null;
    return keadaan.posyandu.filter(function (p) { return rapat(p.nama) === t || rapat(tujuanLengkap(p)) === t; })[0] || null;
  }
  function rapikanTujuan(teks) {
    var p = cariPosyandu(teks);
    return p ? tujuanLengkap(p) : String(teks || "").replace(/\s+/g, " ").trim();
  }
  function sibuk(btn, teks) {
    var asli = btn.innerHTML;
    btn.disabled = true;
    btn.textContent = teks;
    return function () { btn.disabled = false; btn.innerHTML = asli; };
  }

  // ---------------------------------------------------------------- muat data
  function muatData() {
    var s = Simpan.aktif();
    keadaan.status = "memuat";
    tampilStatus();
    var cadangan = function (r) {
      var tugas = [];
      if (r && r.pegawai && r.pegawai.length) keadaan.pegawai = r.pegawai;
      else tugas.push(ambilFile("assets/Template_Input_SPD.xlsx").then(function (buf) {
        keadaan.pegawai = SPDCore.bacaPegawaiResmi(XLSX, buf).map(function (p) {
          return { nama: p.nama, nip: p.nip, pangkat: p.pangkat, jabatan: p.jabatan };
        });
      }));
      if (r && r.posyandu && r.posyandu.length) keadaan.posyandu = r.posyandu;
      else tugas.push(ambilFile("assets/posyandu.json").then(function (buf) {
        keadaan.posyandu = JSON.parse(new TextDecoder().decode(buf));
      }));
      return Promise.all(tugas.map(function (t) { return t.catch(function () { /* tetap lanjut */ }); })).then(function () {
        keadaan.pegawai = keadaan.pegawai.map(function (p) { return Object.assign({}, p, { nama: SPDCore.rapikanNama(p.nama) }); });
      });
    };
    return s.muat().then(function (r) {
      keadaan.pengaturan = Object.assign({}, BAWAAN, r.pengaturan || {});
      keadaan.kegiatan = r.kegiatan || [];
      keadaan.status = r.offline ? "offline" : s.nama;
      keadaan.waktuSalinan = r.waktuSalinan || "";
      return cadangan(r);
    }).catch(function (e) {
      keadaan.status = "gagal";
      keadaan.pesanStatus = e.message;
      return cadangan(null);
    }).then(tampilStatus);
  }

  // Kirim simpanan offline begitu ada internet, lalu muat ulang data terbaru.
  var sinkronBerjalan = null;
  function sinkronkan() {
    if (!Simpan.jumlahAntrian()) return Promise.resolve();
    if (sinkronBerjalan) return sinkronBerjalan;
    sinkronBerjalan = Simpan.sinkron().then(function (r) {
      if (r.terkirim) toast(r.terkirim + " simpanan offline sudah terkirim ke Google Sheet.");
      if (r.catatan.length) alert(r.catatan.join("\n\n"));
      if (r.galat && !r.galat.jaringan) toast("Sinkron gagal: " + r.galat.message, true);
      if (!r.terkirim && !r.catatan.length) { tampilStatus(); return; }
      return muatData().then(function () {
        var h = location.hash;
        if (!keadaan.berubah && !/^#\/(edit|baru)/.test(h)) rute();
      });
    }).then(function () { sinkronBerjalan = null; }, function () { sinkronBerjalan = null; });
    return sinkronBerjalan;
  }

  function tampilStatus() {
    var el = $("status-simpan"), b = $("banner"), antri = Simpan.jumlahAntrian();
    var teksAntri = antri ? '<span class="lencana lencana-antri">' + antri + " simpanan menunggu sinkron</span>" : "";
    b.hidden = true;
    b.className = "banner";
    if (keadaan.status === "memuat") el.innerHTML = '<span class="titik"></span>Memuat data…';
    else if (keadaan.status === "sheets") el.innerHTML = '<b><span class="titik ok"></span>Tersimpan di Google Sheet</b>Data bisa dibuka dari perangkat mana saja.' + teksAntri;
    else if (keadaan.status === "offline") {
      el.innerHTML = '<b><span class="titik lokal"></span>Offline</b>Simpanan dikirim ke Google Sheet saat internet tersambung.' + teksAntri;
      b.hidden = false;
      b.innerHTML = "Sedang offline. Menampilkan data terakhir dari Google Sheet" + (keadaan.waktuSalinan ? " (" + esc(keadaan.waktuSalinan) + ")" : "") +
        ". Kegiatan tetap bisa dibuat, disimpan, dan diunduh; simpanan dikirim otomatis saat internet tersambung.";
    } else if (keadaan.status === "lokal") {
      el.innerHTML = '<b><span class="titik lokal"></span>Tersimpan di perangkat ini</b>Hubungkan Google Sheet di Pengaturan.';
      b.hidden = false;
      b.innerHTML = 'Google Sheet belum terhubung, jadi data hanya tersimpan di browser perangkat ini. <a href="#/pengaturan">Atur koneksi</a>.';
    } else {
      el.innerHTML = '<b><span class="titik gagal"></span>Tidak terhubung</b>' + esc(keadaan.pesanStatus) + teksAntri;
      b.hidden = false;
      b.className = "banner gagal";
      b.innerHTML = esc(keadaan.pesanStatus) + ' <a href="#/pengaturan">Buka Pengaturan</a>.';
    }
  }

  // ---------------------------------------------------------------- router
  var HALAMAN = {
    dashboard: { judul: "Dashboard", tampil: tampilDashboard },
    baru: { judul: "Buat Kegiatan Baru", tampil: function () { bukaForm(null); } },
    edit: { judul: "Ubah Kegiatan", menu: "kegiatan", tampil: function (id) { bukaForm(id); } },
    kegiatan: { judul: "Daftar Kegiatan", tampil: tampilDaftar },
    pegawai: { judul: "Data Pegawai", tampil: tampilPegawai },
    posyandu: { judul: "Data Posyandu", tampil: tampilPosyandu },
    pengaturan: { judul: "Pengaturan", tampil: tampilPengaturan }
  };
  function rute() {
    var bagian = location.hash.replace(/^#\/?/, "").split("/");
    var nama = HALAMAN[bagian[0]] ? bagian[0] : "dashboard";
    var h = HALAMAN[nama];
    $("judul-halaman").textContent = h.judul;
    document.title = h.judul + " · SPD Moncongloe";
    Array.prototype.forEach.call(document.querySelectorAll(".menu-item"), function (a) {
      a.classList.toggle("aktif", a.getAttribute("data-menu") === (h.menu || nama));
    });
    document.querySelector(".app").classList.remove("menu-buka");
    window.scrollTo(0, 0);
    h.tampil(decodeURIComponent(bagian[1] || ""));
  }
  window.addEventListener("hashchange", function () {
    if (keadaan.berubah && !confirm("Perubahan pada kegiatan belum disimpan. Tinggalkan halaman ini?")) {
      history.back();
      return;
    }
    keadaan.berubah = false;
    rute();
  });
  window.addEventListener("beforeunload", function (e) { if (keadaan.berubah) { e.preventDefault(); e.returnValue = ""; } });
  $("btn-menu").addEventListener("click", function () { document.querySelector(".app").classList.add("menu-buka"); });
  $("tirai").addEventListener("click", function () { document.querySelector(".app").classList.remove("menu-buka"); });

  // ---------------------------------------------------------------- dashboard
  function tampilDashboard() {
    var tahun = new Date().getFullYear(), ini = keadaan.kegiatan.filter(function (k) { return String(k.tglKeluar).slice(0, 4) === String(tahun); });
    var kunj = ini.reduce(function (a, k) { return a + (+k.jumlahKunjungan || 0); }, 0);
    var terbaru = urutKegiatan(keadaan.kegiatan).slice(0, 6);
    $("isi").innerHTML =
      '<div class="statistik">' +
      stat(ini.length, "Kegiatan tahun " + tahun) + stat(kunj, "Kunjungan tahun " + tahun) +
      stat(keadaan.kegiatan.length, "Semua kegiatan tersimpan") + stat(keadaan.pegawai.length, "Pegawai") +
      '</div>' +
      '<div class="kartu"><div class="kartu-kepala"><h2>Kegiatan terbaru</h2>' +
      '<a class="tombol" href="#/baru">+ Kegiatan Baru</a></div>' + tabelKegiatan(terbaru) + '</div>';
    pasangAksiDaftar();
  }
  function stat(n, label) { return '<div class="stat"><strong>' + esc(n) + '</strong><span>' + esc(label) + '</span></div>'; }
  function urutKegiatan(list) {
    return list.slice().sort(function (a, b) {
      var x = String(b.tglKeluar || ""), y = String(a.tglKeluar || "");
      return x < y ? -1 : x > y ? 1 : String(b.diubah || "").localeCompare(String(a.diubah || ""));
    });
  }
  function tabelKegiatan(list) {
    if (!list.length) {
      return '<div class="wadah-tabel"><table><tr><td class="kosong">Belum ada kegiatan. ' +
        '<a href="#/baru">Buat kegiatan pertama</a>.</td></tr></table></div>';
    }
    return '<div class="wadah-tabel"><table><thead><tr><th>Tanggal ST</th><th>Kegiatan</th><th>No. Surat Tugas</th>' +
      '<th class="angka">Pegawai</th><th class="angka">Kunjungan</th><th></th></tr></thead><tbody>' +
      list.map(function (k) {
        return '<tr><td>' + esc(tglTampil(k.tglKeluar)) + '</td><td><a href="#/edit/' + encodeURIComponent(k.id) + '">' +
          esc(k.maksud || "(tanpa nama)") + '</a><div class="redup">' + (k.tertunda ? '<span class="lencana">Belum tersinkron</span>' : "Diubah " + esc(k.diubah || "")) + '</div></td>' +
          '<td>' + esc(k.nomorST || "—") + '</td><td class="angka">' + esc(k.jumlahPegawai) + '</td>' +
          '<td class="angka">' + esc(k.jumlahKunjungan) + '</td><td class="aksi">' +
          '<a class="tombol tombol-hantu tombol-kecil" href="#/edit/' + encodeURIComponent(k.id) + '">Buka</a> ' +
          '<span class="menu-unduh"><button class="tombol tombol-hantu tombol-kecil" data-unduh-id="' + esc(k.id) + '">Unduh ▾</button></span> ' +
          '<button class="tombol tombol-bahaya tombol-kecil" data-hapus-id="' + esc(k.id) + '">Hapus</button></td></tr>';
      }).join("") + '</tbody></table></div>';
  }
  function pasangAksiDaftar() {
    Array.prototype.forEach.call(document.querySelectorAll("[data-unduh-id]"), function (btn) {
      btn.addEventListener("click", function (e) {
        e.stopPropagation();
        bukaMenuUnduh(btn, true, function (jenis, tombol) {
          var selesai = sibuk(tombol, "Menyiapkan…");
          Simpan.aktif().ambil(btn.getAttribute("data-unduh-id"))
            .then(function (r) { return buatKeluaran(jenis, r.kegiatan, r.penugasan); })
            .catch(function (e) { toast(e.message, true); })
            .then(selesai);
        });
      });
    });
    Array.prototype.forEach.call(document.querySelectorAll("[data-hapus-id]"), function (btn) {
      btn.addEventListener("click", function () {
        var id = btn.getAttribute("data-hapus-id");
        var k = keadaan.kegiatan.filter(function (x) { return x.id === id; })[0];
        if (!confirm("Hapus kegiatan \"" + (k && k.maksud) + "\" beserta seluruh penugasannya? Tindakan ini tidak bisa dibatalkan.")) return;
        var selesai = sibuk(btn, "Menghapus…");
        Simpan.aktif().hapusKegiatan(id).then(function () {
          keadaan.kegiatan = keadaan.kegiatan.filter(function (x) { return x.id !== id; });
          toast("Kegiatan dihapus.");
          rute();
        }).catch(function (e) { selesai(); toast(e.message, true); });
      });
    });
  }

  // Menu kecil berisi empat jenis file keluaran.
  function bukaMenuUnduh(btn, keBawah, pilih) {
    tutupMenuUnduh();
    var wadah = btn.parentNode, daftar = document.createElement("div");
    daftar.className = "daftar-unduh" + (keBawah ? " ke-bawah" : "");
    Object.keys(KELUARAN).forEach(function (j) {
      var b = document.createElement("button");
      b.type = "button";
      b.textContent = KELUARAN[j].label;
      b.addEventListener("click", function (e) { e.stopPropagation(); tutupMenuUnduh(); pilih(j, btn); });
      daftar.appendChild(b);
    });
    var semua = document.createElement("button");
    semua.type = "button";
    semua.innerHTML = "<b>Semua (4 file)</b>";
    semua.addEventListener("click", function (e) { e.stopPropagation(); tutupMenuUnduh(); pilih("semua", btn); });
    daftar.appendChild(semua);
    wadah.appendChild(daftar);
  }
  function tutupMenuUnduh() {
    Array.prototype.forEach.call(document.querySelectorAll(".daftar-unduh"), function (d) { d.remove(); });
  }
  document.addEventListener("click", tutupMenuUnduh);

  // ---------------------------------------------------------------- daftar kegiatan
  function tampilDaftar() {
    var tahun = {};
    keadaan.kegiatan.forEach(function (k) { if (k.tglKeluar) tahun[String(k.tglKeluar).slice(0, 4)] = 1; });
    $("isi").innerHTML = '<div class="kartu"><div class="kartu-kepala">' +
      '<div class="baris-tombol"><input id="cari-kegiatan" class="cari" type="search" placeholder="Cari kegiatan atau nomor ST…">' +
      '<select id="filter-bulan" style="width:auto"><option value="">Semua bulan</option>' +
      Object.keys(tahun).sort().reverse().map(function (t) {
        return BULAN.map(function (b, i) { return '<option value="' + t + "-" + pad2(i + 1) + '">' + b + " " + t + '</option>'; }).reverse().join("");
      }).join("") + '</select></div>' +
      '<a class="tombol" href="#/baru">+ Kegiatan Baru</a></div><div id="wadah-daftar"></div></div>';
    var saring = function () {
      var q = $("cari-kegiatan").value.toLowerCase(), bln = $("filter-bulan").value;
      var list = urutKegiatan(keadaan.kegiatan).filter(function (k) {
        return (!q || (String(k.maksud) + " " + String(k.nomorST)).toLowerCase().indexOf(q) >= 0) &&
          (!bln || String(k.tglKeluar).slice(0, 7) === bln);
      });
      $("wadah-daftar").innerHTML = tabelKegiatan(list);
      pasangAksiDaftar();
    };
    $("cari-kegiatan").addEventListener("input", saring);
    $("filter-bulan").addEventListener("change", saring);
    saring();
  }

  // ---------------------------------------------------------------- form kegiatan
  function formBaru() {
    var k = {};
    KUNCI_PENGATURAN.forEach(function (x) { k[x] = keadaan.pengaturan[x] || ""; });
    k.id = ""; k.nomorST = ""; k.maksud = ""; k.tglKeluar = hariIni(); k.tglSPD = "";
    return { kegiatan: k, penugasan: [baris(), baris(), baris()] };
  }
  function baris(asal) { return Object.assign({ nomor: "", nama: "", tujuan: "", tanggal: "", tglSPD: "" }, asal || {}); }

  function bukaForm(id) {
    keadaan.berubah = false;
    if (!id) { keadaan.form = formBaru(); tampilForm(); return; }
    $("isi").innerHTML = '<div class="kartu kosong">Memuat kegiatan…</div>';
    Simpan.aktif().ambil(id).then(function (r) {
      var k = Object.assign({}, r.kegiatan);
      keadaan.form = { kegiatan: k, penugasan: r.penugasan.map(baris) };
      if (!keadaan.form.penugasan.length) keadaan.form.penugasan.push(baris());
      tampilForm();
    }).catch(function (e) {
      $("isi").innerHTML = '<div class="kotak kotak-galat">' + esc(e.message) + '</div>';
    });
  }

  function isian(def, nilai) {
    var k = def[0], label = def[1], jenis = def[2], v = esc(nilai == null ? "" : nilai);
    var kontrol = jenis === "textarea"
      ? '<textarea data-f="' + k + '" rows="2">' + v + '</textarea>'
      : '<input data-f="' + k + '" type="' + jenis + '" value="' + v + '"' + (jenis === "number" ? ' min="0" step="1000"' : "") + '>';
    return '<label class="isian' + (jenis === "textarea" ? " lebar" : "") + '">' + esc(label) + kontrol + '</label>';
  }
  function bagian(judul, ket, defs, k, buka) {
    return '<details class="bagian"' + (buka ? " open" : "") + '><summary><span>' + esc(judul) +
      '<span class="redup">' + esc(ket) + '</span></span></summary><div class="isi-bagian"><div class="grid-isian">' +
      defs.map(function (d) { return isian(d, k[d[0]]); }).join("") + '</div></div></details>';
  }

  function tampilForm() {
    var f = keadaan.form, k = f.kegiatan;
    $("isi").innerHTML =
      '<div id="hasil-periksa"></div>' +
      bagian("Surat Tugas", "", BAGIAN_ST, k, true) +
      '<div class="kartu"><div class="kartu-kepala"><div><h2>Penugasan</h2>' +
      '<div class="redup">Satu baris = satu kunjungan. Pegawai yang sama memakai Nomor SPD yang sama.</div></div>' +
      '<div class="baris-tombol"><input id="no-awal" type="number" placeholder="No SPD awal" style="width:130px">' +
      '<button class="tombol tombol-hantu tombol-kecil" id="btn-nomor">Beri nomor otomatis</button></div></div>' +
      '<div class="wadah-tabel"><table class="tabel-tugas"><thead><tr><th>#</th><th class="kol-no">No SPD</th><th>Pegawai</th>' +
      '<th>Tempat tujuan</th><th class="kol-tgl">Tanggal kunjungan</th>' +
      '<th class="kol-tgl" title="Kosongkan bila sama dengan Tanggal SPD di atas">Tgl SPD <span class="redup">(bila beda)</span></th><th></th></tr></thead><tbody id="isi-tugas"></tbody></table></div>' +
      '<div class="baris-tombol" style="margin-top:10px"><button class="tombol tombol-hantu tombol-kecil" id="btn-tambah">+ Tambah baris</button>' +
      '<button class="tombol tombol-hantu tombol-kecil" id="btn-tambah5">+ 5 baris</button>' +
      '<span class="redup" id="ringkas-tugas"></span></div>' +
      '<datalist id="dl-pegawai">' + keadaan.pegawai.map(function (p) { return '<option value="' + esc(p.nama) + '">'; }).join("") + '</datalist>' +
      '<datalist id="dl-tujuan"></datalist></div>' +
      bagian("Rincian SPD", "terisi dari Pengaturan", BAGIAN_SPD, k, false) +
      bagian("Penanda tangan", "Kepala Puskesmas & PPTK", BAGIAN_TTD, k, false) +
      bagian("Anggaran", "untuk Lembar Penerimaan & Verifikasi", BAGIAN_ANGGARAN, k, false) +
      bagian("Verifikator", "untuk Lembar Verifikasi", BAGIAN_VERIF, k, false) +
      '<div class="bilah-aksi">' +
      '<button class="tombol tombol-hantu" id="btn-impor">Impor dari Excel</button>' +
      '<a class="tautan redup" href="assets/Template_Input_SPD.xlsx" download>Unduh template Excel</a>' +
      '<span class="pemisah"></span>' +
      '<button class="tombol tombol-hantu" id="btn-periksa">Periksa</button>' +
      '<span class="menu-unduh"><button class="tombol tombol-hantu" id="btn-unduh">Unduh ▴</button></span>' +
      '<button class="tombol" id="btn-simpan">Simpan</button></div>';

    Array.prototype.forEach.call(document.querySelectorAll("[data-f]"), function (el) {
      el.addEventListener("input", function () { k[el.getAttribute("data-f")] = el.value; keadaan.berubah = true; });
    });
    tampilTugas();
    $("btn-tambah").addEventListener("click", function () { tambahBaris(1); });
    $("btn-tambah5").addEventListener("click", function () { tambahBaris(5); });
    $("btn-nomor").addEventListener("click", nomorOtomatis);
    $("btn-periksa").addEventListener("click", function () { periksaForm(true); });
    $("btn-simpan").addEventListener("click", simpanForm);
    $("btn-impor").addEventListener("click", function () { $("input-impor").value = ""; $("input-impor").click(); });
    $("btn-unduh").addEventListener("click", function (e) {
      e.stopPropagation();
      var btn = this;
      bukaMenuUnduh(btn, false, function (jenis) {
        var h = periksaForm(false);
        if (h.galat.length) { toast("Perbaiki dulu kesalahan yang ditandai merah.", true); return; }
        var selesai = sibuk(btn, "Menyiapkan…");
        buatKeluaran(jenis, f.kegiatan, f.penugasan).catch(function (e) { toast(e.message, true); }).then(selesai);
      });
    });
  }

  function tampilTugas() {
    var f = keadaan.form, peta = petaPegawai();
    $("isi-tugas").innerHTML = f.penugasan.map(function (r, i) {
      var p = peta[kunciNama(r.nama)];
      var kelas = (r.nama && !p ? " salah" : "") + (r.tujuan && !cariPosyandu(r.tujuan) ? " salah-tujuan" : "");
      return '<tr data-i="' + i + '"' + (kelas ? ' class="' + kelas.trim() + '"' : "") + '><td class="redup">' + (i + 1) + '</td>' +
        '<td><input data-k="nomor" value="' + esc(r.nomor) + '" inputmode="numeric"></td>' +
        '<td><input data-k="nama" list="dl-pegawai" value="' + esc(r.nama) + '" placeholder="Ketik nama…">' +
        '<div class="jabatan">' + esc(p ? p.jabatan : (r.nama ? "Tidak ada di Data Pegawai" : "")) + '</div></td>' +
        '<td><input data-k="tujuan" list="dl-tujuan" value="' + esc(r.tujuan) + '" placeholder="Ketik nama posyandu…">' +
        '<div class="jabatan">' + esc(ketTujuan(r.tujuan)) + '</div></td>' +
        '<td><input data-k="tanggal" type="date" value="' + esc(r.tanggal) + '"></td>' +
        '<td><input data-k="tglSPD" type="date" value="' + esc(r.tglSPD || "") + '" title="Kosongkan bila sama dengan Tanggal SPD di atas"></td>' +
        '<td class="aksi"><button class="tombol tombol-hantu tombol-kecil" data-salin="' + i + '" title="Salin baris">Salin</button> ' +
        '<button class="tombol tombol-bahaya tombol-kecil" data-hapus="' + i + '" title="Hapus baris">✕</button></td></tr>';
    }).join("");
    $("dl-tujuan").innerHTML = keadaan.posyandu.map(function (p) { return '<option value="' + esc(tujuanLengkap(p)) + '">'; }).join("");
    Array.prototype.forEach.call(document.querySelectorAll("#isi-tugas input"), function (el) {
      el.addEventListener("input", function () {
        var i = +el.closest("tr").getAttribute("data-i");
        f.penugasan[i][el.getAttribute("data-k")] = el.value;
        keadaan.berubah = true;
        if (el.getAttribute("data-k") === "nama") {
          var p = petaPegawai()[kunciNama(el.value)];
          el.parentNode.querySelector(".jabatan").textContent = p ? p.jabatan : (el.value ? "Tidak ada di Data Pegawai" : "");
          el.closest("tr").classList.toggle("salah", !!el.value && !p);
        }
        if (el.getAttribute("data-k") === "tujuan") tandaiTujuan(el);
        ringkasTugas();
      });
      if (el.getAttribute("data-k") === "tujuan") {
        el.addEventListener("change", function () {
          var rapi = rapikanTujuan(el.value);
          if (rapi !== el.value) { el.value = rapi; f.penugasan[+el.closest("tr").getAttribute("data-i")].tujuan = rapi; }
          tandaiTujuan(el);
        });
      }
    });
    Array.prototype.forEach.call(document.querySelectorAll("[data-hapus]"), function (b) {
      b.addEventListener("click", function () {
        f.penugasan.splice(+b.getAttribute("data-hapus"), 1);
        if (!f.penugasan.length) f.penugasan.push(baris());
        keadaan.berubah = true;
        tampilTugas();
      });
    });
    Array.prototype.forEach.call(document.querySelectorAll("[data-salin]"), function (b) {
      b.addEventListener("click", function () {
        var i = +b.getAttribute("data-salin");
        f.penugasan.splice(i + 1, 0, baris({ nomor: f.penugasan[i].nomor, nama: f.penugasan[i].nama, tujuan: "", tanggal: "", tglSPD: f.penugasan[i].tglSPD || "" }));
        keadaan.berubah = true;
        tampilTugas();
      });
    });
    ringkasTugas();
  }
  function ketTujuan(t) {
    if (!t) return "";
    var p = cariPosyandu(t);
    return p ? "Desa " + p.desa : "Tidak ada di Data Posyandu";
  }
  function tandaiTujuan(el) {
    el.parentNode.querySelector(".jabatan").textContent = ketTujuan(el.value);
    el.closest("tr").classList.toggle("salah-tujuan", !!el.value.trim() && !cariPosyandu(el.value));
  }
  function ringkasTugas() {
    var isi = keadaan.form.penugasan.filter(function (r) { return r.nama || r.tujuan || r.tanggal; });
    var peg = {};
    isi.forEach(function (r) { if (r.nama) peg[r.nama.toLowerCase()] = 1; });
    $("ringkas-tugas").textContent = isi.length + " kunjungan · " + Object.keys(peg).length + " pegawai";
  }
  function tambahBaris(n) {
    for (var i = 0; i < n; i++) keadaan.form.penugasan.push(baris());
    tampilTugas();
  }
  function nomorOtomatis() {
    var awal = parseInt($("no-awal").value, 10);
    if (!awal) { toast("Isi dulu Nomor SPD awal, misalnya 2307.", true); $("no-awal").focus(); return; }
    var nomorPeg = {}, n = awal;
    keadaan.form.penugasan.forEach(function (r) {
      if (!r.nama) return;
      var key = r.nama.trim().toLowerCase();
      if (!(key in nomorPeg)) nomorPeg[key] = n++;
      r.nomor = String(nomorPeg[key]);
    });
    keadaan.berubah = true;
    tampilTugas();
    toast("Nomor SPD " + awal + "–" + (n - 1) + " dibagikan ke " + (n - awal) + " pegawai.");
  }

  function penugasanTerisi() {
    return keadaan.form.penugasan.filter(function (r) { return (r.nomor + r.nama + r.tujuan + r.tanggal).trim(); });
  }
  function periksaRekaman(kegiatan, penugasan) {
    var isi = [];
    penugasan.forEach(function (r, i) {
      if ((String(r.nomor || "") + (r.nama || "") + (r.tujuan || "") + (r.tanggal || "")).trim()) isi.push(Object.assign({ baris: i + 1 }, r));
    });
    var data = SPDCore.dariRekaman(kegiatan, isi.map(function (r) {
      return Object.assign({}, r, { tujuan: rapikanTujuan(r.tujuan) });
    }), keadaan.pegawai);
    var h = SPDCore.periksa(data, null);
    isi.forEach(function (r) {
      if (r.tujuan && !cariPosyandu(r.tujuan)) {
        h.peringatan.push("Penugasan baris " + r.baris + ": '" + r.tujuan + "' tidak ada di Data Posyandu, jadi kolom Desa di Surat Tugas berisi '-'.");
      }
    });
    // Pesan dari pemeriksa memakai istilah template Excel; sesuaikan dengan form.
    var ganti = function (s) {
      return s.replace(/^Sheet Kegiatan: /, "").replace(/^Penugasan baris (\d+)/, "Penugasan baris $1")
        .replace(/tidak ada di sheet Pegawai/, "tidak ada di Data Pegawai").replace(/Sheet Pegawai: /, "Data Pegawai: ")
        .replace(/ Unduh template terbaru bila baris ini belum ada\./, "")
        .replace(/^Sheet Penugasan belum berisi kunjungan\./, "Penugasan belum berisi kunjungan.");
    };
    h.galat = h.galat.map(ganti);
    h.peringatan = h.peringatan.map(ganti);
    return h;
  }
  function periksaForm(tampilkanOk) {
    var h = periksaRekaman(keadaan.form.kegiatan, keadaan.form.penugasan), el = $("hasil-periksa"), html = "";
    if (h.galat.length) html += '<div class="kotak kotak-galat"><b>Perlu diperbaiki</b><ul>' + h.galat.map(function (s) { return "<li>" + esc(s) + "</li>"; }).join("") + "</ul></div>";
    if (h.peringatan.length) html += '<div class="kotak kotak-peringatan"><b>Perhatikan</b><ul>' + h.peringatan.map(function (s) { return "<li>" + esc(s) + "</li>"; }).join("") + "</ul></div>";
    if (!h.galat.length && tampilkanOk) {
      var kunj = h.spd.reduce(function (a, s) { return a + s.kunjungan.length; }, 0);
      html += '<div class="kotak kotak-ok"><b>Data lengkap.</b> ' + h.spd.length + " SPD, " + kunj + " kunjungan. Siap diunduh.</div>";
    }
    el.innerHTML = html;
    if (html && (h.galat.length || tampilkanOk)) el.scrollIntoView({ behavior: "smooth", block: "start" });
    return h;
  }
  function simpanForm() {
    var btn = $("btn-simpan"), f = keadaan.form;
    if (!f.kegiatan.maksud) { toast("Isi dulu nama kegiatan.", true); return; }
    f.penugasan.forEach(function (r) { r.tujuan = rapikanTujuan(r.tujuan); });
    var h = periksaForm(false);
    var selesai = sibuk(btn, "Menyimpan…");
    var penugasan = penugasanTerisi();
    var simpan = function (paksa) { return Simpan.aktif().simpanKegiatan(f.kegiatan, penugasan, paksa); };
    simpan(false).catch(function (e) {
      if (e.kode !== "KONFLIK") throw e;
      if (!confirm(e.message + "\n\nTimpa dengan isian Anda? Pilih Batal untuk membuka versi terbaru (isian Anda tidak disimpan).")) {
        keadaan.berubah = false;
        bukaForm(f.kegiatan.id);
        throw new Error("Versi terbaru dibuka.");
      }
      return simpan(true);
    }).then(function (r) {
      var baru = !f.kegiatan.id;
      f.kegiatan.id = r.id;
      if (!r.tertunda) f.kegiatan.diubah = r.diubah;
      keadaan.berubah = false;
      var peg = {};
      penugasan.forEach(function (x) { peg[x.nama.toLowerCase()] = 1; });
      var ring = { id: r.id, nomorST: f.kegiatan.nomorST, tglKeluar: f.kegiatan.tglKeluar, maksud: f.kegiatan.maksud,
        diubah: r.diubah, tertunda: !!r.tertunda, jumlahKunjungan: penugasan.length, jumlahPegawai: Object.keys(peg).length };
      keadaan.kegiatan = keadaan.kegiatan.filter(function (x) { return x.id !== r.id; }).concat([ring]);
      toast(r.tertunda ? "Tersimpan di perangkat. Akan dikirim ke Google Sheet saat internet tersambung."
        : h.galat.length ? "Tersimpan sebagai draf. Masih ada data yang perlu diperbaiki." : "Kegiatan tersimpan.");
      if (baru) history.replaceState(null, "", "#/edit/" + encodeURIComponent(r.id));
      $("judul-halaman").textContent = "Ubah Kegiatan";
      tampilTugas();
      tampilStatus();
    }).catch(function (e) { toast(e.message, true); }).then(selesai);
  }

  function buatKeluaran(jenis, kegiatan, penugasan) {
    var h = periksaRekaman(kegiatan, penugasan);
    if (h.galat.length) return Promise.reject(new Error("Data kegiatan belum lengkap: " + h.galat[0]));
    var daftar = jenis === "semua" ? Object.keys(KELUARAN) : [jenis];
    return daftar.reduce(function (p, j) {
      return p.then(function () {
        var o = KELUARAN[j];
        return ambilFile(o.master)
          .then(function (m) { return SPDCore[o.buat](env, m, h); })
          .then(function (r) { unduh(r.data, r.namaFile, o.tipe); });
      });
    }, Promise.resolve()).then(function () {
      toast(daftar.length > 1 ? "Empat file sudah diunduh." : KELUARAN[jenis].label + " sudah diunduh.");
    });
  }

  // Impor template Excel lama ke form.
  $("input-impor").addEventListener("change", function () {
    var file = this.files[0];
    if (!file || !keadaan.form) return;
    file.arrayBuffer().then(function (buf) {
      var d = SPDCore.bacaTemplate(XLSX, buf), k = keadaan.form.kegiatan;
      Object.keys(d.kegiatan).forEach(function (x) {
        if (x === "tglKeluarRaw" || x === "tglKeluar") return;
        if (d.kegiatan[x]) k[x] = d.kegiatan[x];
      });
      k.tglKeluar = isoDari(d.kegiatan.tglKeluarRaw) || k.tglKeluar;
      var normal = {};
      keadaan.pegawai.forEach(function (p) { normal[kunciNama(p.nama)] = p.nama; });
      keadaan.form.penugasan = d.penugasan.map(function (r) {
        return baris({ nomor: r.nomor.replace(/\.0+$/, ""), nama: normal[kunciNama(r.nama)] || r.nama, tujuan: rapikanTujuan(r.tujuan), tanggal: isoDari(r.tanggalRaw) });
      });
      if (!keadaan.form.penugasan.length) keadaan.form.penugasan.push(baris());
      keadaan.berubah = true;
      tampilForm();
      periksaForm(true);
      toast("Data dari " + file.name + " dimasukkan ke form. Periksa lalu Simpan.");
    }).catch(function (e) { toast(e.message || String(e), true); });
  });

  // ---------------------------------------------------------------- pegawai & posyandu
  // Tabel induk yang bisa diubah langsung: cari, tambah, hapus, simpan.
  function tabelInduk(o) {
    var data = o.data().map(function (x) { var y = {}; o.kolom.forEach(function (c) { y[c[0]] = x[c[0]] || ""; }); return y; });
    $("isi").innerHTML = '<div class="kartu"><div class="kartu-kepala"><div class="baris-tombol">' +
      '<input id="cari-induk" class="cari" type="search" placeholder="' + esc(o.cari) + '">' +
      '<span class="redup" id="jumlah-induk"></span></div><div class="baris-tombol">' +
      '<button class="tombol tombol-hantu" id="btn-tambah-induk">+ Tambah ' + esc(o.satuan) + '</button>' +
      '<button class="tombol" id="btn-simpan-induk">Simpan</button></div></div>' +
      (o.ket ? '<p class="redup" style="margin-top:0">' + o.ket + '</p>' : "") +
      '<div class="wadah-tabel"><table class="tabel-tugas"><thead><tr>' +
      o.kolom.map(function (c) { return "<th>" + esc(c[1]) + "</th>"; }).join("") +
      '<th></th></tr></thead><tbody id="isi-induk"></tbody></table></div></div>';
    var gambar = function () {
      var q = $("cari-induk").value.toLowerCase(), tampil = 0;
      $("isi-induk").innerHTML = data.map(function (x, i) {
        var cocok = !q || o.kolom.map(function (c) { return x[c[0]]; }).join(" ").toLowerCase().indexOf(q) >= 0;
        if (cocok) tampil++;
        return '<tr data-i="' + i + '"' + (cocok ? "" : " hidden") + '>' +
          o.kolom.map(function (c) { return '<td><input data-c="' + c[0] + '" value="' + esc(x[c[0]]) + '"></td>'; }).join("") +
          '<td class="aksi"><button class="tombol tombol-bahaya tombol-kecil" data-hapus-induk="' + i + '">✕</button></td></tr>';
      }).join("");
      $("jumlah-induk").textContent = tampil + " dari " + data.length + " " + o.satuan;
      Array.prototype.forEach.call(document.querySelectorAll("#isi-induk input"), function (el) {
        el.addEventListener("input", function () { data[+el.closest("tr").getAttribute("data-i")][el.getAttribute("data-c")] = el.value; });
      });
      Array.prototype.forEach.call(document.querySelectorAll("[data-hapus-induk]"), function (b) {
        b.addEventListener("click", function () {
          var x = data[+b.getAttribute("data-hapus-induk")];
          if (!confirm("Hapus " + (x.nama || "baris ini") + " dari daftar " + o.satuan + "?")) return;
          data.splice(+b.getAttribute("data-hapus-induk"), 1);
          gambar();
        });
      });
    };
    $("cari-induk").addEventListener("input", gambar);
    $("btn-tambah-induk").addEventListener("click", function () {
      $("cari-induk").value = "";
      data.unshift(o.baru());
      gambar();
      document.querySelector('#isi-induk input[data-c="nama"]').focus();
    });
    $("btn-simpan-induk").addEventListener("click", function () {
      var btn = this, bersih = data.filter(function (x) { return String(x.nama).trim(); }).map(o.bersihkan);
      var ganda = {}, dobel = [];
      bersih.forEach(function (x) { var k = x.nama.toLowerCase(); if (ganda[k]) dobel.push(x.nama); ganda[k] = 1; });
      if (dobel.length) { toast("Nama ganda: " + dobel.join(", ") + ". Bedakan dulu namanya.", true); return; }
      var salah = o.periksa ? o.periksa(bersih) : "";
      if (salah) { toast(salah, true); return; }
      var selesai = sibuk(btn, "Menyimpan…");
      o.simpan(bersih).then(function (r) {
        toast((r.tertunda ? "Tersimpan di perangkat, dikirim saat internet tersambung" : "Data " + o.satuan + " tersimpan") + " (" + r.jumlah + " " + o.satuan + ").");
        tampilStatus();
        o.tampil();
      }).catch(function (e) { selesai(); toast(e.message, true); });
    });
    gambar();
  }
  function tampilPegawai() {
    tabelInduk({
      data: function () { return keadaan.pegawai.map(function (p) { return Object.assign({}, p, { nama: SPDCore.rapikanNama(p.nama) }); }); }, satuan: "pegawai", cari: "Cari nama, NIP, atau jabatan…",
      kolom: [["nama", "Nama (dan gelar)"], ["nip", "NIP / NIK"], ["pangkat", "Pangkat / Gol."], ["jabatan", "Jabatan"]],
      baru: function () { return { nama: "", nip: "", pangkat: "-", jabatan: "" }; },
      bersihkan: function (p) {
        return { nama: SPDCore.rapikanNama(p.nama), nip: String(p.nip || "").replace(/\s+/g, ""), pangkat: String(p.pangkat || "").trim() || "-", jabatan: String(p.jabatan || "").trim() || "-" };
      },
      simpan: function (bersih) {
        return Simpan.aktif().simpanPegawai(bersih).then(function (r) {
          keadaan.pegawai = r.pegawai || bersih;
          return { tertunda: r.tertunda, jumlah: keadaan.pegawai.length };
        });
      },
      tampil: tampilPegawai
    });
  }
  function tampilPosyandu() {
    tabelInduk({
      data: function () { return keadaan.posyandu; }, satuan: "posyandu", cari: "Cari posyandu atau desa…",
      ket: "Dipakai untuk pilihan Tempat Tujuan. Di Surat Tugas, Tempat Kegiatan ditulis &ldquo;Posyandu Desa …&rdquo; dan kolom Desa terisi otomatis.",
      kolom: [["nama", "Nama Posyandu"], ["desa", "Desa"]],
      baru: function () { return { nama: "", desa: "" }; },
      bersihkan: function (p) { return { nama: String(p.nama).replace(/\s+/g, " ").trim(), desa: String(p.desa || "").replace(/\s+/g, " ").trim() }; },
      periksa: function (list) {
        var kosong = list.filter(function (p) { return !p.desa; }).map(function (p) { return p.nama; });
        return kosong.length ? "Desa belum diisi untuk: " + kosong.join(", ") + "." : "";
      },
      simpan: function (bersih) {
        return Simpan.aktif().simpanPosyandu(bersih).then(function (r) {
          keadaan.posyandu = r.posyandu || bersih;
          return { tertunda: r.tertunda, jumlah: keadaan.posyandu.length };
        });
      },
      tampil: tampilPosyandu
    });
  }

  // ---------------------------------------------------------------- pengaturan
  function tampilPengaturan() {
    var p = keadaan.pengaturan, dariConfig = window.SPD_CONFIG && window.SPD_CONFIG.apiUrl;
    $("isi").innerHTML =
      '<div class="kartu"><h2>Koneksi Google Sheet</h2>' +
      '<p class="redup">Isi URL Web App dari Google Sheet "Database SPD Moncongloe" dan PIN yang tertulis di sheet Pengaturan. ' +
      'Kosongkan URL untuk menyimpan di perangkat ini saja.</p>' +
      '<div class="grid-isian"><label class="isian lebar">URL Web App (berakhiran /exec)<input id="in-url" type="url" value="' + esc(Simpan.urlApi()) + '" placeholder="https://script.google.com/macros/s/…/exec"></label>' +
      '<label class="isian">PIN<input id="in-pin" type="password" value="' + esc(Simpan.pin()) + '" autocomplete="off"></label></div>' +
      (dariConfig ? '<p class="redup">URL bawaan sudah diatur di js/config.js untuk semua petugas.</p>' : "") +
      '<div class="baris-tombol" style="margin-top:12px"><button class="tombol" id="btn-hubung">Simpan &amp; Hubungkan</button>' +
      (Simpan.adaDataLokal() ? '<button class="tombol tombol-hantu" id="btn-pindah">Pindahkan data perangkat ke Google Sheet</button>' : "") +
      '</div><div id="hasil-hubung" style="margin-top:12px"></div></div>' +
      '<div class="kartu"><div class="kartu-kepala"><div><h2>Nilai bawaan form kegiatan</h2>' +
      '<div class="redup">Dipakai otomatis setiap membuat kegiatan baru. Bisa diubah per kegiatan.</div></div>' +
      '<button class="tombol" id="btn-simpan-pengaturan">Simpan Pengaturan</button></div>' +
      '<div class="grid-isian">' + SEMUA_BAGIAN.filter(function (d) { return KUNCI_PENGATURAN.indexOf(d[0]) >= 0; })
        .map(function (d) { return isian(d, p[d[0]]); }).join("") + '</div></div>';

    $("btn-hubung").addEventListener("click", function () {
      var btn = this, url = $("in-url").value.trim(), pin = $("in-pin").value.trim();
      if (url && !/^https:\/\/script\.google\.com\/.+\/exec$/.test(url)) {
        $("hasil-hubung").innerHTML = '<div class="kotak kotak-galat">URL harus berbentuk https://script.google.com/macros/s/…/exec</div>';
        return;
      }
      Simpan.atur(url === dariConfig ? "" : url, pin);
      var selesai = sibuk(btn, "Menghubungkan…");
      muatData().then(function () {
        selesai();
        var ok = keadaan.status === "sheets" || keadaan.status === "lokal";
        $("hasil-hubung").innerHTML = ok
          ? '<div class="kotak kotak-ok">' + (keadaan.status === "sheets" ? "Terhubung ke Google Sheet. " + keadaan.kegiatan.length + " kegiatan, " + keadaan.pegawai.length + " pegawai." : "Data disimpan di perangkat ini.") + "</div>"
          : '<div class="kotak kotak-galat">' + esc(keadaan.pesanStatus) + "</div>";
      });
    });
    var pindah = $("btn-pindah");
    if (pindah) pindah.addEventListener("click", function () { pindahkanLokal(this); });
    $("btn-simpan-pengaturan").addEventListener("click", function () {
      var btn = this, nilai = {};
      Array.prototype.forEach.call(document.querySelectorAll("[data-f]"), function (el) { nilai[el.getAttribute("data-f")] = el.value; });
      ["pejabatNama", "pptkNama", "verifKelengkapan", "verifPersiapan", "verifPptkUkm", "verifBendahara", "verifPjKeuangan"]
        .forEach(function (x) { if (nilai[x]) nilai[x] = SPDCore.rapikanNama(nilai[x]); });
      var selesai = sibuk(btn, "Menyimpan…");
      Simpan.aktif().simpanPengaturan(nilai).then(function () {
        Object.assign(keadaan.pengaturan, nilai);
        toast("Pengaturan tersimpan.");
      }).catch(function (e) { toast(e.message, true); }).then(selesai);
    });
  }

  // Data yang sempat tersimpan di browser dipindahkan ke Google Sheet.
  function pindahkanLokal(btn) {
    if (Simpan.aktif().nama !== "sheets" || keadaan.status !== "sheets") { toast("Hubungkan Google Sheet dulu.", true); return; }
    var lokal = JSON.parse(localStorage.getItem("spd_data_lokal") || "{}"), daftar = lokal.kegiatan || [];
    if (!confirm("Pindahkan " + daftar.length + " kegiatan dari perangkat ini ke Google Sheet?")) return;
    var selesai = sibuk(btn, "Memindahkan…");
    daftar.reduce(function (p, k) {
      return p.then(function () { return Simpan.aktif().simpanKegiatan(Object.assign({}, k, { id: "" }), (lokal.penugasan || {})[k.id] || []); });
    }, Promise.resolve()).then(function () {
      localStorage.removeItem("spd_data_lokal");
      toast(daftar.length + " kegiatan dipindahkan ke Google Sheet.");
      return muatData();
    }).catch(function (e) { toast(e.message, true); }).then(function () { selesai(); tampilPengaturan(); });
  }

  // ---------------------------------------------------------------- PWA
  var promptPasang = null;
  window.addEventListener("beforeinstallprompt", function (e) { e.preventDefault(); promptPasang = e; $("btn-pasang").hidden = false; });
  $("btn-pasang").addEventListener("click", function () {
    if (!promptPasang) return;
    promptPasang.prompt();
    promptPasang.userChoice.then(function () { promptPasang = null; $("btn-pasang").hidden = true; });
  });
  function statusJaringan() { $("status-jaringan").hidden = navigator.onLine; }
  window.addEventListener("online", function () {
    statusJaringan();
    if (keadaan.status === "offline" || keadaan.status === "gagal") {
      sinkronkan().then(function () { if (keadaan.status !== "sheets") return muatData(); }).then(function () {
        if (!keadaan.berubah && !/^#\/(edit|baru)/.test(location.hash)) rute();
      });
    } else sinkronkan();
  });
  window.addEventListener("offline", function () {
    statusJaringan();
    if (keadaan.status === "sheets") { keadaan.status = "offline"; tampilStatus(); }
  });
  statusJaringan();
  if ("serviceWorker" in navigator && location.protocol !== "file:") {
    navigator.serviceWorker.register("sw.js").then(function (reg) {
      reg.addEventListener("updatefound", function () {
        var baru = reg.installing;
        baru && baru.addEventListener("statechange", function () {
          if (baru.state === "installed" && navigator.serviceWorker.controller) toast("Versi baru aplikasi tersedia. Muat ulang halaman untuk memakainya.");
        });
      });
    }).catch(function () { /* tanpa mode offline */ });
  }

  // ---------------------------------------------------------------- mulai
  $("isi").innerHTML = '<div class="kartu kosong">Memuat data…</div>';
  muatData().then(rute).then(function () { if (keadaan.status === "sheets") return sinkronkan(); });
  setInterval(function () { if (navigator.onLine && Simpan.jumlahAntrian()) sinkronkan(); }, 60000);
})();
