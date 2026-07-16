// Caches the whole app shell on install so the PWA keeps working with no network —
// your data lives in localStorage on-device regardless, this is just about the app
// itself (HTML/JS/fonts) still loading when you're offline.
// IMPORTANT: bump this version string on every deploy — cache-first means an
// installed phone keeps serving the old app shell forever otherwise.
const CACHE_NAME = "arbor-animals-v15";
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

// index.html is the entire app (HTML/CSS/JS all inline, ~3MB single file) — it's
// the one thing that changes on every deploy, and it's exactly what kept getting
// stuck on stale cached copies despite bumping CACHE_NAME: a device already
// controlled by an old worker doesn't reliably re-fetch it until some future
// visit's background update check happens to land. Serve it network-first
// instead, falling back to cache only when offline, so a fresh deploy reaches an
// already-installed device on its very next launch — no manual cache-clear or
// second relaunch needed. Static assets (fonts/vendor libs/icons) change rarely,
// so they stay cache-first for speed and offline use.
self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  // Never intercept the GitHub API — Gist reads (cross-device sync pulls) are GET
  // requests, and without this exclusion they fell into the cache-first branch
  // below just like any other same-origin asset: the FIRST successful fetch to a
  // given gist URL got cached forever, so every later pull silently kept
  // returning that one frozen snapshot no matter what actually changed on
  // GitHub. Confirmed live: 5 fetches to api.github.com/zen (which normally
  // returns a random quote per request) all returned the identical first
  // response. API calls must always hit the network live — a Gist read failing
  // should surface as a real error, not silently serve stale data.
  if (new URL(event.request.url).hostname === "api.github.com") return;
  const isAppShell = event.request.mode === "navigate"
    || event.request.url.endsWith("/index.html")
    || event.request.url.endsWith("/Arbor/")
    || event.request.url.endsWith("/Arbor");
  if (isAppShell) {
    event.respondWith(
      fetch(event.request).then((res) => {
        if (res.ok) {
          const clone = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        }
        return res;
      }).catch(() => caches.match(event.request))
    );
    return;
  }
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
