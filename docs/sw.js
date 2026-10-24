// Lưu sẵn toàn bộ app trên máy để mở được khi không có mạng. Đổi VERSION mỗi lần sửa app để máy tải bản mới.
const VERSION = 'aegis-mobile-v3';
const FILES = ['./', 'index.html', 'style.css', 'app.js', 'core.js', 'vendor/xlsx.full.min.js', 'config.json',
               'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);

  // config.json: có mạng thì lấy bản mới nhất ngay (để sửa cấu hình là điện thoại thấy), mất mạng thì dùng bản đã lưu
  if (url.pathname.endsWith('/config.json')) {
    e.respondWith(fetch(e.request).then(r => {
      if (r.ok) caches.open(VERSION).then(c => c.put(e.request, r.clone()));
      return r;
    }).catch(() => caches.match(e.request, { ignoreSearch: true })));
    return;
  }

  // Còn lại: ưu tiên bản đã lưu; có mạng thì âm thầm cập nhật cho lần mở sau
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then(hit => {
    const net = fetch(e.request).then(r => {
      if (r.ok) caches.open(VERSION).then(c => c.put(e.request, r.clone()));
      return r;
    }).catch(() => hit);
    return hit || net;
  }));
});
