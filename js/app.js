/* Antarmuka Pembuat SPD — menghubungkan halaman dengan SPDCore. */
(function () {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };
  var env = { JSZip: window.JSZip, DOMParser: window.DOMParser, XMLSerializer: window.XMLSerializer };
  var keadaan = { hasil: null, namaSheet: [] };

  // ---------------------------------------------------------------- util
  function el(tag, attr, anak) {
    var e = document.createElement(tag);
    Object.keys(attr || {}).forEach(function (k) {
      if (k === "text") e.textContent = attr[k];
      else if (k === "class") e.className = attr[k];
      else e.setAttribute(k, attr[k]);
    });
    (anak || []).forEach(function (a) { if (a) e.appendChild(a); });
    return e;
  }
  function toast(teks) {
    var t = $("toast");
    t.textContent = teks;
    t.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { t.hidden = true; }, 4000);
  }
  function ambil(url) {
    return fetch(url, { cache: "no-cache" }).then(function (r) {
      if (!r.ok) throw new Error("Gagal memuat " + url + " (" + r.status + ")");
      return r.arrayBuffer();
    });
  }
  function isiDaftar(ul, items) {
    ul.innerHTML = "";
    items.forEach(function (s) { ul.appendChild(el("li", { text: s })); });
  }

  // ---------------------------------------------------------------- unggah
  var zona = $("zona-unggah"), input = $("input-file");
  input.addEventListener("change", function () { if (input.files[0]) proses(input.files[0]); });
  ["dragenter", "dragover"].forEach(function (ev) {
    zona.addEventListener(ev, function (e) { e.preventDefault(); zona.classList.add("aktif"); });
  });
  ["dragleave", "drop"].forEach(function (ev) {
    zona.addEventListener(ev, function (e) { e.preventDefault(); zona.classList.remove("aktif"); });
  });
  zona.addEventListener("drop", function (e) {
    var f = e.dataTransfer.files[0];
    if (f) proses(f);
  });

  function proses(file) {
    $("nama-file").textContent = file.name;
    $("pesan-buat").textContent = "";
    if (!/\.xlsx$/i.test(file.name)) {
      tampilkanGagal("File harus berformat .xlsx (Excel). Jika memakai .xls lama, simpan ulang sebagai .xlsx.");
      return;
    }
    Promise.all([
      file.arrayBuffer(),
      ambil("assets/Template_Input_SPD.xlsx").catch(function () { return null; })
    ]).then(function (hasil) {
      var data = SPDCore.bacaTemplate(XLSX, hasil[0]);
      var resmi = hasil[1] ? SPDCore.bacaPegawaiResmi(XLSX, hasil[1]) : null;
      keadaan.hasil = SPDCore.periksa(data, resmi);
      tampilkan(keadaan.hasil);
    }).catch(function (e) {
      tampilkanGagal(e.message || String(e));
    });
  }

  function tampilkanGagal(pesan) {
    keadaan.hasil = null;
    $("langkah-3").hidden = false;
    $("langkah-4").hidden = true;
    $("ringkasan").innerHTML = "";
    $("kotak-peringatan").hidden = true;
    $("wadah-tabel").hidden = true;
    $("kotak-galat").hidden = false;
    isiDaftar($("daftar-galat"), [pesan]);
  }

  // ---------------------------------------------------------------- tampilan hasil
  function tampilkan(h) {
    $("langkah-3").hidden = false;

    var pegawai = {}, kunjungan = 0;
    h.spd.forEach(function (s) { pegawai[s.pegawai.nama] = 1; kunjungan += s.kunjungan.length; });
    var ring = $("ringkasan");
    ring.innerHTML = "";
    [[h.spd.length, "lembar SPD"], [Object.keys(pegawai).length, "pegawai"], [kunjungan, "kunjungan"]]
      .forEach(function (x) {
        ring.appendChild(el("div", { class: "angka-ringkas" }, [
          el("strong", { text: String(x[0]) }), el("span", { text: x[1] })
        ]));
      });
    if (h.kegiatan.maksud) {
      ring.appendChild(el("p", { class: "kegiatan", text: h.kegiatan.maksud + (h.kegiatan.nomorST ? " · ST " + h.kegiatan.nomorST : "") }));
    }

    $("kotak-galat").hidden = !h.galat.length;
    isiDaftar($("daftar-galat"), h.galat);
    $("kotak-peringatan").hidden = !h.peringatan.length;
    isiDaftar($("daftar-peringatan"), h.peringatan);

    var tbody = $("isi-tabel");
    tbody.innerHTML = "";
    keadaan.namaSheet = h.spd.length && !h.galat.length ? SPDCore.pratinjauNamaSheet(h.spd) : [];
    h.spd.forEach(function (s, i) {
      var selNomor;
      if (s.nomor) {
        selNomor = el("td", { class: "nomor", text: s.nomor });
      } else {
        var inp = el("input", {
          type: "text", class: "input-nomor", "data-kunci": s.kunci,
          placeholder: "Nomor SPD tambahan", "aria-label": "Nomor SPD tambahan untuk " + s.pegawai.nama
        });
        inp.addEventListener("input", cekSiap);
        selNomor = el("td", { class: "nomor" }, [inp, el("small", {
          text: "Bagian " + s.bagian + " dari " + s.jumlahBagian + " (lanjutan " + s.nomorAsal + ")"
        })]);
      }
      tbody.appendChild(el("tr", s.nomor ? {} : { class: "perlu-nomor" }, [
        el("td", { class: "sheet", text: keadaan.namaSheet[i] || "—" }),
        selNomor,
        el("td", { class: "pegawai", text: s.pegawai.nama }),
        el("td", { class: "angka", text: String(s.kunjungan.length) }),
        el("td", { class: "tgl", text: SPDCore.daftarTanggal(s.kunjungan.map(function (v) { return v.tgl; })) })
      ]));
    });
    $("wadah-tabel").hidden = !h.spd.length;

    $("langkah-4").hidden = !!h.galat.length;
    cekSiap();
    $("langkah-3").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function nomorTambahan() {
    var out = {};
    Array.prototype.forEach.call(document.querySelectorAll(".input-nomor"), function (i) {
      out[i.getAttribute("data-kunci")] = i.value.trim();
    });
    return out;
  }

  function cekSiap() {
    var btn = $("btn-buat"), kosong = 0;
    var tambahan = nomorTambahan();
    Object.keys(tambahan).forEach(function (k) { if (!tambahan[k]) kosong++; });
    btn.disabled = !keadaan.hasil || keadaan.hasil.galat.length > 0 || kosong > 0;
    $("pesan-buat").textContent = kosong ? "Isi " + kosong + " Nomor SPD tambahan di tabel atas terlebih dahulu." : "";
    $("pesan-buat").className = "pesan";
  }

  // ---------------------------------------------------------------- buat file
  $("btn-buat").addEventListener("click", function () {
    var btn = this, pesan = $("pesan-buat");
    btn.disabled = true;
    pesan.textContent = "Menyusun file SPD…";
    pesan.className = "pesan";
    ambil("assets/master_spd.xlsx")
      .then(function (master) { return SPDCore.buatSPD(env, master, keadaan.hasil, nomorTambahan()); })
      .then(function (r) {
        var blob = new Blob([r.data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        var a = el("a", { href: URL.createObjectURL(blob), download: r.namaFile });
        document.body.appendChild(a);
        a.click();
        setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
        pesan.textContent = "Selesai: " + r.namaFile + " (" + r.sheet.length + " sheet) sudah diunduh.";
        pesan.className = "pesan sukses";
      })
      .catch(function (e) {
        pesan.textContent = "Gagal: " + (e.message || e);
        pesan.className = "pesan gagal";
      })
      .then(function () { btn.disabled = false; });
  });

  // ---------------------------------------------------------------- PWA
  var promptPasang = null;
  window.addEventListener("beforeinstallprompt", function (e) {
    e.preventDefault();
    promptPasang = e;
    $("btn-pasang").hidden = false;
  });
  $("btn-pasang").addEventListener("click", function () {
    if (!promptPasang) return;
    promptPasang.prompt();
    promptPasang.userChoice.then(function () { promptPasang = null; $("btn-pasang").hidden = true; });
  });

  function statusJaringan() { $("status-jaringan").hidden = navigator.onLine; }
  window.addEventListener("online", statusJaringan);
  window.addEventListener("offline", statusJaringan);
  statusJaringan();

  if ("serviceWorker" in navigator && location.protocol !== "file:") {
    navigator.serviceWorker.register("sw.js").then(function (reg) {
      reg.addEventListener("updatefound", function () {
        var baru = reg.installing;
        baru && baru.addEventListener("statechange", function () {
          if (baru.state === "installed" && navigator.serviceWorker.controller) {
            toast("Versi baru aplikasi tersedia. Muat ulang halaman untuk memakainya.");
          }
        });
      });
    }).catch(function () { /* tanpa mode offline */ });
  }
})();
