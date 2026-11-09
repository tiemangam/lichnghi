// Gặm Quản lý nghỉ – service worker: giữ khung app + hiện thông báo; dữ liệu lấy trực tiếp từ Google
var V = 'gam-quanly-v6';
var SHELL = ['./', './index.html', './manifest.webmanifest', '../icon-ql-192.png', '../logo.png'];
self.addEventListener('install', function (e) { e.waitUntil(caches.open(V).then(function (c) { return c.addAll(SHELL); })); self.skipWaiting(); });
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (ks) { return Promise.all(ks.filter(function (k) { return k !== V && k.indexOf('gam-quanly') === 0; }).map(function (k) { return caches.delete(k); })); }));
  self.clients.claim();
});
self.addEventListener('fetch', function (e) {
  var u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  // luôn hỏi lại GitHub (bỏ qua bộ nhớ đệm 10 phút của trình duyệt) để sửa app là máy nhân viên thấy ngay; mất mạng thì dùng bản đã lưu
  e.respondWith(fetch(e.request.url, { cache: 'no-cache', credentials: 'same-origin', redirect: e.request.mode === 'navigate' ? 'manual' : 'follow' }).then(function (r) {
    if (r.ok) { var copy = r.clone(); caches.open(V).then(function (c) { c.put(e.request, copy); }); }
    return r;
  }).catch(function () { return caches.match(e.request, { ignoreSearch: true }); }));
});
self.addEventListener('notificationclick', function (e) {
  e.notification.close();
  e.waitUntil(self.clients.matchAll({ type: 'window' }).then(function (cs) {
    for (var i = 0; i < cs.length; i++) if ('focus' in cs[i]) return cs[i].focus();
    return self.clients.openWindow('./?mo=1');
  }));
});
