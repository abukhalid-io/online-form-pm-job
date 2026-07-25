// Service worker PM Control Panel
// Tugasnya cuma cache "app shell" (HTML/CSS/JS/ikon) supaya app kebuka
// instan meski jaringan lambat/offline. Data (fetch ke Apps Script) SELALU
// diambil langsung dari network, tidak pernah di-cache, supaya data yang
// ditampilkan selalu yang terbaru.

var CACHE_NAME = 'pm-control-panel-shell-v1';
var SHELL_FILES = [
  '/',
  '/index.html',
  '/manifest.json',
  '/icon-192.png',
  '/icon-512.png'
];

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) {
      return cache.addAll(SHELL_FILES);
    }).then(function () {
      return self.skipWaiting();
    })
  );
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(
        keys.filter(function (key) { return key !== CACHE_NAME; })
          .map(function (key) { return caches.delete(key); })
      );
    }).then(function () {
      return self.clients.claim();
    })
  );
});

self.addEventListener('fetch', function (event) {
  var url = event.request.url;

  // Jangan pernah cache request ke Apps Script (data harus selalu fresh).
  if (url.indexOf('script.google.com') > -1 || url.indexOf('script.googleusercontent.com') > -1) {
    event.respondWith(fetch(event.request));
    return;
  }

  // App shell: cache-first, fallback ke network kalau belum ke-cache.
  event.respondWith(
    caches.match(event.request).then(function (cached) {
      return cached || fetch(event.request).then(function (response) {
        var copy = response.clone();
        caches.open(CACHE_NAME).then(function (cache) { cache.put(event.request, copy); });
        return response;
      });
    }).catch(function () {
      return caches.match('/index.html');
    })
  );
});
