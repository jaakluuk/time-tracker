// Serves the app from the phone's cache instantly, then quietly fetches any update
// for next time. Requests to Google (your data) are never cached.
var CACHE = "time-tracker-v1";
var ASSETS = ["./", "./index.html", "./manifest.webmanifest", "./icon-192.png", "./icon-512.png", "./apple-touch-icon.png"];

self.addEventListener("install", function(e) {
  e.waitUntil(caches.open(CACHE).then(function(c) { return c.addAll(ASSETS); }));
  self.skipWaiting();
});

self.addEventListener("activate", function(e) {
  e.waitUntil(caches.keys().then(function(keys) {
    return Promise.all(keys.filter(function(k) { return k !== CACHE; }).map(function(k) { return caches.delete(k); }));
  }).then(function() { return self.clients.claim(); }));
});

self.addEventListener("fetch", function(e) {
  var url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== self.location.origin) return;

  e.respondWith(caches.open(CACHE).then(function(cache) {
    return cache.match(e.request, { ignoreSearch: true }).then(function(hit) {
      var network = fetch(e.request).then(function(res) {
        if (res && res.ok) cache.put(e.request, res.clone());
        return res;
      }).catch(function() { return hit; });
      if (hit) { e.waitUntil(network); return hit; }
      return network;
    });
  }));
});
