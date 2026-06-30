// Service worker - cho phép cài như app & chạy giao diện offline
const CACHE = 'tksx-v15';
const ASSETS = [
  './', './index.html', './kiemke.html', './manifest.json', './icon-192.png', './icon-512.png',
  'https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js'
];
self.addEventListener('install', e => { self.skipWaiting(); e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS).catch(()=>{}))); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k))))); self.clients.claim(); });
self.addEventListener('fetch', e => {
  const url = e.request.url;
  // Không cache lệnh gọi Google (tải DS / gửi dữ liệu) để luôn lấy mới
  if (url.indexOf('script.google') !== -1) return;
  e.respondWith(
    caches.match(e.request).then(r => r || fetch(e.request).then(resp => {
      if (e.request.method === 'GET' && resp.ok) { const c = resp.clone(); caches.open(CACHE).then(ca => ca.put(e.request, c)); }
      return resp;
    }).catch(() => caches.match('./index.html')))
  );
});
