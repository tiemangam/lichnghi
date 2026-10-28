// Nhà Gặm (trước là Gặm Lịch nghỉ) – service worker: giữ sẵn khung app + quy trình để mở nhanh, đọc được cả khi mất mạng; lịch nghỉ vẫn lấy trực tiếp từ Google
var V = 'gam-lichnghi-v7';
var SHELL = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png', './logo.png', './quytrinh.js', './img/qt2.png', './img/qt5.png', './img/qt7.png'];
self.addEventListener('install', function (e) { e.waitUntil(caches.open(V).then(function (c) { return c.addAll(SHELL); })); self.skipWaiting(); });
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (ks) { return Promise.all(ks.filter(function (k) { return k !== V && k.indexOf('gam-lichnghi') === 0; }).map(function (k) { return caches.delete(k); })); }));
  self.clients.claim();
});
self.addEventListener('fetch', function (e) {
  var u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  e.respondWith(fetch(e.request).then(function (r) {
    var copy = r.clone(); caches.open(V).then(function (c) { c.put(e.request, copy); }); return r;
  }).catch(function () { return caches.match(e.request, { ignoreSearch: true }); }));
});
