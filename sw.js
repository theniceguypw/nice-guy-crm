// Keeps the app shell available offline. Data always comes from the network
// (Supabase and Microsoft), so the app never shows stale customer records.
const CACHE = "nice-guy-crm-v1";
const SHELL = ["./", "index.html", "styles.css", "app.js", "config.js", "manifest.json", "icons/icon-192.png", "icons/icon-512.png", "icons/apple-touch-icon.png"];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE).then((cache) => cache.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  // Only handle our own files. Let Supabase, Microsoft and CDN requests go straight to the network.
  if (event.request.method !== "GET" || url.origin !== self.location.origin) return;

  // Network first, fall back to the cached copy when offline.
  event.respondWith(
    fetch(event.request)
      .then((response) => {
        const copy = response.clone();
        caches.open(CACHE).then((cache) => cache.put(event.request, copy));
        return response;
      })
      .catch(() => caches.match(event.request).then((hit) => hit || caches.match("index.html")))
  );
});
