/* Service worker: aplikasi tetap bisa dipakai tanpa internet.
 * Strategi "jaringan dulu": saat online selalu mengambil versi terbaru dari
 * GitHub (termasuk template & daftar pegawai), saat offline memakai salinan
 * terakhir yang tersimpan. Tidak perlu mengubah file ini saat memperbarui
 * template atau kode. */
const CACHE = "spd-moncongloe";
const INTI = [
  "./", "index.html", "css/app.css", "js/app.js", "js/spd-core.js",
  "js/vendor/xlsx.mini.min.js", "js/vendor/jszip.min.js",
  "assets/Template_Input_SPD.xlsx", "assets/master_spd.xlsx",
  "contoh/Contoh_Input_Kelas_Ibu_Balita_Agustus_2026.xlsx",
  "manifest.webmanifest", "icons/icon-192.png", "icons/icon-512.png"
];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(INTI)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => { e.waitUntil(self.clients.claim()); });

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET" || new URL(req.url).origin !== location.origin) return;
  e.respondWith(
    fetch(req, { cache: "no-cache" })
      .then((res) => {
        if (res.ok) {
          const salinan = res.clone();
          caches.open(CACHE).then((c) => c.put(req, salinan));
        }
        return res;
      })
      .catch(() => caches.match(req, { ignoreSearch: true }).then((r) => r || caches.match("index.html")))
  );
});
