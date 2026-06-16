const CACHE = "bhdc-tracker-v2";
const ASSETS = [
  "./",
  "./index.html",
  "./bhdc-logo.png",
  "./manifest.json",
];

// Install — cache core assets
self.addEventListener("install", e => {
  e.waitUntil(
    caches.open(CACHE).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

// Activate — delete any old cache versions
self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Fetch strategy:
// - Google Apps Script calls → always go to network (never cache)
// - Everything else → cache first, fall back to network
self.addEventListener("fetch", e => {
  if (
    e.request.method === "POST" ||
    e.request.url.includes("script.google.com") ||
    e.request.url.includes("fonts.googleapis.com") ||
    e.request.url.includes("fonts.gstatic.com")
  ) {
    // Network only — never intercept these
    return;
  }

  e.respondWith(
    caches.match(e.request).then(cached => {
      // Return cached version, but also fetch fresh in background
      const networkFetch = fetch(e.request).then(response => {
        if (response && response.status === 200) {
          const clone = response.clone();
          caches.open(CACHE).then(cache => cache.put(e.request, clone));
        }
        return response;
      }).catch(() => cached); // if offline, fall back to cache

      return cached || networkFetch;
    })
  );
});
