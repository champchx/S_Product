// เก็บหน้าแอปไว้ในเครื่อง เปิดเร็วและเปิดได้แม้สัญญาณอ่อน ส่วนข้อมูลจาก Google Sheets ไม่ถูกเก็บ (ดึงสดเสมอ)
const CACHE = 'bscm-rice-v3';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png', './logo-white.png', './logo-black.png', './logo-blue.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== 'GET' || u.origin !== location.origin) return; // API / ฟอนต์ ไม่แตะ
  // ลองโหลดจากเน็ตก่อน เพื่อให้ได้หน้าแอปเวอร์ชันล่าสุดเสมอ ถ้าไม่มีเน็ตค่อยใช้ของที่เก็บไว้
  e.respondWith(fetch(r).then(res => {
    if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(r, copy)); }
    return res;
  }).catch(() => caches.match(r).then(m => m || caches.match('./index.html'))));
});
