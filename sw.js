// Service worker: يخلّي التطبيق يفتح حتى بدون إنترنت.
// الشبكة أولاً (عشان أي تحديث يوصل فوراً)، والنسخة المحفوظة بس إذا ما في نت.
const CACHE = "azi-bots-v2";

self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil(self.clients.claim()));

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET" || !req.url.startsWith("http")) return;
  // ملفات الموقع نفسه: نتأكد من السيرفر كل مرة بدل ما ناخذ نسخة المتصفح القديمة (max-age=600)
  const own = new URL(req.url).origin === self.location.origin;
  e.respondWith(
    (own ? fetch(req.url, { cache: "no-cache" }) : fetch(req))
      .then(res => {
        if (res.ok || res.type === "opaque") {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(req, copy));
        }
        return res;
      })
      .catch(() => caches.match(req, { ignoreSearch: true }))
  );
});
