/*
 * Penyimpanan data aplikasi. Dua cara dengan antarmuka yang sama:
 *  - Sheets: Google Sheet lewat Web App Apps Script (apps-script/Code.gs).
 *    Saat offline, data terakhir dibaca dari salinan di perangkat dan
 *    simpanan masuk antrian yang dikirim otomatis begitu internet tersambung.
 *  - Lokal : localStorage browser, dipakai bila Google Sheet belum diatur.
 * Semua fungsi mengembalikan Promise.
 */
(function (root) {
  "use strict";

  var KUNCI_URL = "spd_api_url", KUNCI_PIN = "spd_pin", KUNCI_LOKAL = "spd_data_lokal";
  var KUNCI_SALINAN = "spd_salinan", KUNCI_ANTRIAN = "spd_antrian", KUNCI_PETA_ID = "spd_peta_id";
  var MAKS_DETAIL = 40;

  function baca(k) { try { return localStorage.getItem(k) || ""; } catch (e) { return ""; } }
  function tulis(k, v) { try { if (v) localStorage.setItem(k, v); else localStorage.removeItem(k); } catch (e) { /* abaikan */ } }
  function bacaJson(k, bawaan) { try { return JSON.parse(localStorage.getItem(k)) || bawaan; } catch (e) { return bawaan; } }
  function tulisJson(k, v) {
    try { localStorage.setItem(k, JSON.stringify(v)); }
    catch (e) { throw new Error("Penyimpanan browser penuh atau diblokir."); }
  }

  function urlApi() { return baca(KUNCI_URL) || (root.SPD_CONFIG && root.SPD_CONFIG.apiUrl) || ""; }
  function pin() { return baca(KUNCI_PIN); }
  function sekarang() {
    var d = new Date(), p = function (n) { return (n < 10 ? "0" : "") + n; };
    return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()) + " " + p(d.getHours()) + ":" + p(d.getMinutes());
  }
  function idAcak() { return Math.random().toString(36).slice(2, 10); }
  function urutNama(list) {
    return list.slice().sort(function (a, b) { return String(a.nama).toLowerCase().localeCompare(String(b.nama).toLowerCase()); });
  }
  function ringkas(k, p) {
    var peg = {};
    p.forEach(function (x) { peg[String(x.nama).toLowerCase()] = 1; });
    return { id: k.id, nomorST: k.nomorST, tglKeluar: k.tglKeluar, maksud: k.maksud, diubah: k.diubah,
      jumlahKunjungan: p.length, jumlahPegawai: Object.keys(peg).length };
  }

  // ---------------------------------------------------------------- Google Sheet
  function galatJaringan() {
    var e = new Error("Tidak bisa terhubung ke Google Sheet. Periksa internet dan URL Web App.");
    e.jaringan = true;
    return e;
  }
  function panggil(aksi, data) {
    if (root.navigator && root.navigator.onLine === false) return Promise.reject(galatJaringan());
    var body = Object.assign({ aksi: aksi, pin: pin() }, data || {});
    // Tanpa header Content-Type supaya tidak memicu preflight CORS.
    return fetch(urlApi(), { method: "POST", body: JSON.stringify(body), redirect: "follow" })
      .then(function (r) {
        if (r.status >= 500) throw galatJaringan();
        if (!r.ok) throw new Error("Google Sheet tidak bisa dihubungi (" + r.status + ").");
        return r.json();
      })
      .catch(function (e) {
        if (e instanceof SyntaxError) throw new Error("Jawaban Google Sheet tidak dikenali. Periksa URL Web App.");
        if (e && (e.jaringan || (e.message && e.message.indexOf("Google Sheet") === 0))) throw e;
        throw galatJaringan();
      })
      .then(function (j) {
        if (!j.ok) { var err = new Error(j.pesan || "Gagal."); err.kode = j.kode; throw err; }
        return j;
      });
  }

  // Antrian simpanan offline. Satu entri per kegiatan (yang terbaru), dan satu
  // entri untuk pegawai, posyandu, serta pengaturan.
  function antrian() { return bacaJson(KUNCI_ANTRIAN, []); }
  function petaId() { return bacaJson(KUNCI_PETA_ID, {}); }
  function idAsli(id) { return (id && petaId()[id]) || id; }
  function masukAntrian(entri) {
    var a = antrian(), ganti = false;
    a = a.map(function (x) {
      if (x.aksi !== entri.aksi || (entri.id && x.id !== entri.id)) return x;
      ganti = true;
      // Pertahankan versi dasar saat kegiatan pertama kali diubah offline.
      if (entri.kegiatan && x.kegiatan) entri.kegiatan.diubah = x.kegiatan.diubah;
      return entri;
    });
    if (!ganti) a.push(entri);
    tulisJson(KUNCI_ANTRIAN, a);
  }
  function salinan() { return bacaJson(KUNCI_SALINAN, null); }
  function simpanDetail(id, data) {
    var s = salinan();
    if (!s) return;
    s.detail = s.detail || {};
    s.detail[id] = { kegiatan: data.kegiatan, penugasan: data.penugasan, waktu: Date.now() };
    var ids = Object.keys(s.detail).sort(function (a, b) { return s.detail[b].waktu - s.detail[a].waktu; });
    ids.slice(MAKS_DETAIL).forEach(function (x) { delete s.detail[x]; });
    try { tulisJson(KUNCI_SALINAN, s); } catch (e) { /* salinan hanya cadangan */ }
  }
  // Data dari Google Sheet (atau salinannya) ditimpa dengan isi antrian.
  function denganAntrian(r) {
    var hasil = { pengaturan: r.pengaturan, pegawai: r.pegawai, posyandu: r.posyandu, kegiatan: (r.kegiatan || []).slice() };
    antrian().forEach(function (x) {
      if (x.aksi === "simpanKegiatan") {
        var s = ringkas(x.kegiatan, x.penugasan);
        s.id = x.id;
        s.diubah = "Belum tersinkron";
        s.tertunda = true;
        hasil.kegiatan = hasil.kegiatan.filter(function (k) { return k.id !== x.id; }).concat([s]);
      } else if (x.aksi === "hapusKegiatan") {
        hasil.kegiatan = hasil.kegiatan.filter(function (k) { return k.id !== x.id; });
      } else if (x.aksi === "simpanPegawai") hasil.pegawai = x.pegawai;
      else if (x.aksi === "simpanPosyandu") hasil.posyandu = x.posyandu;
      else if (x.aksi === "simpanPengaturan") hasil.pengaturan = Object.assign({}, hasil.pengaturan, x.pengaturan);
    });
    return hasil;
  }
  function cobaAtauAntri(aksi, data, entri, hasilTertunda) {
    var kirim = antrian().length ? Promise.reject(galatJaringan()) : panggil(aksi, data);
    return kirim.catch(function (e) {
      if (!e.jaringan) throw e;
      masukAntrian(entri);
      return Object.assign({ ok: true, tertunda: true }, hasilTertunda);
    });
  }

  var sedangSinkron = null;
  // Mengirim antrian satu per satu. Berhenti di galat jaringan; sisanya dicoba lagi nanti.
  function sinkron() {
    if (sedangSinkron) return sedangSinkron;
    if (!urlApi() || !antrian().length) return Promise.resolve({ terkirim: 0, catatan: [] });
    var terkirim = 0, catatan = [];
    var langkah = function () {
      var a = antrian();
      if (!a.length) return Promise.resolve();
      var x = a[0], kirim;
      if (x.aksi === "simpanKegiatan") {
        var k = Object.assign({}, x.kegiatan, { id: idAsli(x.id) });
        var baru = /^tmp-/.test(k.id);
        if (baru) k.id = "";
        kirim = panggil("simpanKegiatan", { kegiatan: k, penugasan: x.penugasan }).catch(function (e) {
          if (e.kode !== "KONFLIK") throw e;
          // Sudah diubah orang lain: simpan versi offline sebagai kegiatan terpisah.
          catatan.push("\"" + k.maksud + "\" sudah diubah orang lain, jadi versi dari perangkat ini disimpan sebagai salinan terpisah.");
          return panggil("simpanKegiatan", { kegiatan: Object.assign({}, k, { id: "", maksud: k.maksud + " (salinan offline)" }), penugasan: x.penugasan });
        }).then(function (r) {
          if (baru) { var m = petaId(); m[x.id] = r.id; tulisJson(KUNCI_PETA_ID, m); }
          simpanDetail(r.id, { kegiatan: Object.assign({}, k, { id: r.id, diubah: r.diubah }), penugasan: x.penugasan });
          // Bila kegiatan ini diubah lagi selama dikirim, versi dasarnya ikut diperbarui.
          tulisJson(KUNCI_ANTRIAN, antrian().map(function (y) {
            if (y.aksi === "simpanKegiatan" && y.id === x.id && y.waktu !== x.waktu) y.kegiatan.diubah = r.diubah;
            return y;
          }));
          return r;
        });
      } else if (x.aksi === "hapusKegiatan") {
        kirim = panggil("hapusKegiatan", { id: idAsli(x.id) });
      } else {
        var d = {};
        d[{ simpanPegawai: "pegawai", simpanPosyandu: "posyandu", simpanPengaturan: "pengaturan" }[x.aksi]] =
          x.pegawai || x.posyandu || x.pengaturan;
        kirim = panggil(x.aksi, d);
      }
      return kirim.then(function () {
        tulisJson(KUNCI_ANTRIAN, antrian().filter(function (y) { return !(y.aksi === x.aksi && y.id === x.id && y.waktu === x.waktu); }));
        terkirim++;
        return langkah();
      });
    };
    sedangSinkron = langkah().then(function () {
      return { terkirim: terkirim, catatan: catatan };
    }, function (e) {
      return { terkirim: terkirim, catatan: catatan, galat: e };
    }).then(function (r) { sedangSinkron = null; return r; });
    return sedangSinkron;
  }

  var Sheets = {
    nama: "sheets",
    muat: function () {
      return panggil("muat").then(function (r) {
        var lama = salinan() || {};
        try {
          tulisJson(KUNCI_SALINAN, { pengaturan: r.pengaturan, pegawai: r.pegawai, posyandu: r.posyandu, kegiatan: r.kegiatan,
            detail: lama.detail || {}, waktu: sekarang() });
        } catch (e) { /* salinan hanya cadangan */ }
        return denganAntrian(r);
      }).catch(function (e) {
        var s = salinan();
        if (!e.jaringan || !s) throw e;
        return Object.assign(denganAntrian(s), { offline: true, waktuSalinan: s.waktu });
      });
    },
    ambil: function (id) {
      var x = antrian().filter(function (y) { return y.aksi === "simpanKegiatan" && y.id === id; })[0];
      if (x) return Promise.resolve({ kegiatan: Object.assign({}, x.kegiatan, { id: id }), penugasan: x.penugasan });
      return panggil("ambil", { id: idAsli(id) }).then(function (r) {
        simpanDetail(r.kegiatan.id, r);
        return r;
      }).catch(function (e) {
        var d = (salinan() || {}).detail || {};
        if (!e.jaringan) throw e;
        if (d[idAsli(id)]) return { kegiatan: d[idAsli(id)].kegiatan, penugasan: d[idAsli(id)].penugasan };
        throw new Error("Sedang offline dan kegiatan ini belum pernah dibuka di perangkat ini, jadi isinya belum tersedia.");
      });
    },
    simpanKegiatan: function (k, p, paksa) {
      var id = idAsli(k.id) || "", idAntri = id || "tmp-" + idAcak();
      var kg = Object.assign({}, k, { id: /^tmp-/.test(id) ? "" : id });
      return cobaAtauAntri("simpanKegiatan", { kegiatan: kg, penugasan: p, paksa: !!paksa },
        { aksi: "simpanKegiatan", id: idAntri, kegiatan: Object.assign({}, kg, { id: idAntri }), penugasan: p, waktu: Date.now() },
        { id: idAntri, diubah: k.diubah || "" })
        .then(function (r) {
          if (!r.tertunda) simpanDetail(r.id, { kegiatan: Object.assign({}, kg, { id: r.id, diubah: r.diubah }), penugasan: p });
          return r;
        });
    },
    hapusKegiatan: function (id) {
      if (/^tmp-/.test(id) && !petaId()[id]) {
        tulisJson(KUNCI_ANTRIAN, antrian().filter(function (y) { return y.id !== id; }));
        return Promise.resolve({ ok: true });
      }
      // Simpanan offline kegiatan ini tidak perlu dikirim lagi.
      tulisJson(KUNCI_ANTRIAN, antrian().filter(function (y) { return !(y.aksi === "simpanKegiatan" && y.id === id); }));
      return cobaAtauAntri("hapusKegiatan", { id: idAsli(id) }, { aksi: "hapusKegiatan", id: id, waktu: Date.now() }, {});
    },
    simpanPegawai: function (p) {
      return cobaAtauAntri("simpanPegawai", { pegawai: p }, { aksi: "simpanPegawai", pegawai: urutNama(p), waktu: Date.now() }, { pegawai: urutNama(p) });
    },
    simpanPosyandu: function (p) {
      return cobaAtauAntri("simpanPosyandu", { posyandu: p }, { aksi: "simpanPosyandu", posyandu: urutNama(p), waktu: Date.now() }, { posyandu: urutNama(p) });
    },
    simpanPengaturan: function (p) {
      var a = antrian().filter(function (y) { return y.aksi === "simpanPengaturan"; })[0];
      var gabung = Object.assign({}, a ? a.pengaturan : {}, p);
      return cobaAtauAntri("simpanPengaturan", { pengaturan: p }, { aksi: "simpanPengaturan", pengaturan: gabung, waktu: Date.now() }, {});
    }
  };

  // ---------------------------------------------------------------- lokal
  function dataLokal() { return bacaJson(KUNCI_LOKAL, {}); }
  function simpanLokal(d) { tulisJson(KUNCI_LOKAL, d); }
  var Lokal = {
    nama: "lokal",
    muat: function () {
      return Promise.resolve().then(function () {
        var d = dataLokal();
        return {
          pengaturan: d.pengaturan || null,
          pegawai: d.pegawai || null,
          posyandu: d.posyandu || null,
          kegiatan: (d.kegiatan || []).map(function (k) { return ringkas(k, (d.penugasan || {})[k.id] || []); })
        };
      });
    },
    ambil: function (id) {
      return Promise.resolve().then(function () {
        var d = dataLokal(), k = (d.kegiatan || []).filter(function (x) { return x.id === id; })[0];
        if (!k) throw new Error("Kegiatan tidak ditemukan.");
        return { kegiatan: k, penugasan: (d.penugasan || {})[id] || [] };
      });
    },
    simpanKegiatan: function (k, p) {
      return Promise.resolve().then(function () {
        var d = dataLokal(), daftar = d.kegiatan || [], waktu = sekarang();
        var baru = Object.assign({}, k, { id: k.id || idAcak(), diubah: waktu });
        var ada = false;
        daftar = daftar.map(function (x) {
          if (x.id !== baru.id) return x;
          ada = true; baru.dibuat = x.dibuat; return baru;
        });
        if (!ada) { baru.dibuat = waktu; daftar.push(baru); }
        d.kegiatan = daftar;
        d.penugasan = d.penugasan || {};
        d.penugasan[baru.id] = p;
        simpanLokal(d);
        return { id: baru.id, diubah: waktu };
      });
    },
    hapusKegiatan: function (id) {
      return Promise.resolve().then(function () {
        var d = dataLokal();
        d.kegiatan = (d.kegiatan || []).filter(function (x) { return x.id !== id; });
        if (d.penugasan) delete d.penugasan[id];
        simpanLokal(d);
        return {};
      });
    },
    simpanPegawai: function (p) {
      return Promise.resolve().then(function () {
        var d = dataLokal();
        d.pegawai = urutNama(p);
        simpanLokal(d);
        return { pegawai: d.pegawai };
      });
    },
    simpanPosyandu: function (p) {
      return Promise.resolve().then(function () {
        var d = dataLokal();
        d.posyandu = urutNama(p);
        simpanLokal(d);
        return { posyandu: d.posyandu };
      });
    },
    simpanPengaturan: function (p) {
      return Promise.resolve().then(function () {
        var d = dataLokal();
        d.pengaturan = Object.assign({}, d.pengaturan || {}, p);
        simpanLokal(d);
        return {};
      });
    }
  };

  root.Simpan = {
    aktif: function () { return urlApi() ? Sheets : Lokal; },
    urlApi: urlApi,
    pin: pin,
    atur: function (url, p) {
      if (url !== urlApi()) { tulis(KUNCI_SALINAN, ""); tulis(KUNCI_PETA_ID, ""); }
      tulis(KUNCI_URL, url);
      tulis(KUNCI_PIN, p);
    },
    adaDataLokal: function () { return (dataLokal().kegiatan || []).length > 0; },
    sinkron: sinkron,
    jumlahAntrian: function () { return urlApi() ? antrian().length : 0; }
  };
})(typeof self !== "undefined" ? self : this);
