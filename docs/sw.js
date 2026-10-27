// Lưu sẵn toàn bộ app trên máy để mở được khi không có mạng. Đổi VERSION mỗi lần sửa app để máy tải bản mới.
// Cả bộ file được tải và lưu cùng một lúc (bỏ qua bộ nhớ đệm HTTP) và chỉ đổi khi VERSION đổi, để không bao giờ lẫn
// file cũ với file mới (lẫn là giao diện vỡ).
const VERSION = 'aegis-mobile-v17';
const FILES = ['./', 'index.html', 'style.css', 'app.js', 'core.js', 'i18n.js', 'vendor/xlsx.full.min.js', 'config.json',
               'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION)
    .then(c => Promise.all(FILES.map(f => fetch(new Request(f, { cache: 'reload' })).then(r => { if (!r.ok) throw new Error(f); return c.put(f, r); }))))
    .then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);

  // config.json: có mạng thì lấy bản mới nhất ngay (để sửa cấu hình là điện thoại thấy), mất mạng thì dùng bản đã lưu
  if (url.pathname.endsWith('/config.json')) {
    e.respondWith(fetch(e.request, { cache: 'no-cache' }).then(r => {
      if (r.ok) caches.open(VERSION).then(c => c.put(e.request, r.clone()));
      return r;
    }).catch(() => caches.match(e.request, { ignoreSearch: true })));
    return;
  }

  // Còn lại: chỉ lấy từ bộ file đã lưu của phiên bản này; không có thì mới ra mạng
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then(hit => hit || fetch(e.request)));
});
