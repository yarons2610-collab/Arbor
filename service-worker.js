// Caches the whole app shell on install so the PWA keeps working with no network —
// your data lives in localStorage on-device regardless, this is just about the app
// itself (HTML/JS/fonts) still loading when you're offline.
// IMPORTANT: bump this version string on every deploy — cache-first means an
// installed phone keeps serving the old app shell forever otherwise.
const CACHE_NAME = "arbor-animals-v5";
const APP_SHELL = [
  "./",
  "index.html",
  "manifest.json",
  "vendor/d3.min.js",
  "vendor/three.min.js",
  "vendor/fonts.css",
  "vendor/fonts/jetbrains-mono-latin.woff2",
  "vendor/fonts/space-grotesk-latin.woff2",
  "icons/icon-192.png",
  "icons/icon-512.png",
  "icons/apple-touch-icon.png",
  "icons/favicon-32.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// cache-first for the app shell, falling back to network (and caching the result)
// for anything not pre-cached
self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request).then((res) => {
        if (res.ok) {
          const clone = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        }
        return res;
      }).catch(() => cached);
    })
  );
});
