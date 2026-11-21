/* เจ้าคุณผจญภัย: เก็บไฟล์เกมไว้ในเครื่องให้เล่นได้ตอนไม่มีเน็ต
   หน้าเกม: ลองโหลดจากเน็ตก่อน (จะได้รุ่นใหม่เสมอ) ถ้าไม่มีเน็ตจึงใช้ที่เก็บไว้
   ไฟล์อื่น (ไอคอน ฯลฯ): ใช้ที่เก็บไว้ก่อน */
const CACHE = 'chaokhun-v2';
const FILES = ['./', 'index.html', 'manifest.json', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/apple-touch-icon.png', 'icons/favicon-32.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES.map(f => new Request(f, { cache: 'reload' })))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  const isPage = req.mode === 'navigate' || req.url.endsWith('index.html');
  if (isPage) {
    e.respondWith(
      /* cache: 'no-cache' = ถามเซิร์ฟเวอร์ทุกครั้งว่ามีรุ่นใหม่ไหม ไม่ใช้ของที่เบราว์เซอร์จำไว้ */
      fetch(req.url, { cache: 'no-cache' }).then(res => { if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put('./', copy)); } return res; })
        .catch(() => caches.match('./').then(r => r || caches.match('index.html')))
    );
  } else {
    e.respondWith(caches.match(req).then(r => r || fetch(req).then(res => { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); return res; })));
  }
});
