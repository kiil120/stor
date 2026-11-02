/* ذاكرة مؤقتة عند العميل: الصور والخطوط والملفات الثابتة تُحمَّل من جهازه فوراً وتُحدَّث في الخلفية */
const C = 'dafi-v2';
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(
  caches.keys().then(k => Promise.all(k.filter(x => x !== C).map(x => caches.delete(x)))).then(() => self.clients.claim())));
self.addEventListener('fetch', e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== 'GET' || u.hostname.endsWith('script.google.com') || u.hostname === 'script.googleusercontent.com') return;
  if (u.pathname.endsWith('admin.html') || u.pathname.endsWith('sw.js')) return;
  if (r.mode === 'navigate') { // الصفحة: الشبكة أولاً ثم النسخة المحفوظة عند انقطاع الإنترنت
    e.respondWith(fetch(r).then(res => { const c = res.clone(); caches.open(C).then(x => x.put(r, c)); return res; }).catch(() => caches.match(r)));
    return;
  }
  const ok = u.origin === location.origin || r.destination === 'image' || u.hostname.indexOf('fonts.g') === 0;
  if (!ok) return;
  e.respondWith(caches.open(C).then(async c => { // الباقي: من الجهاز فوراً ثم تحديث بالخلفية
    const hit = await c.match(r);
    const net = fetch(r).then(res => { if (res && (res.ok || res.type === 'opaque')) c.put(r, res.clone()); return res; }).catch(() => hit);
    return hit || net;
  }));
});
