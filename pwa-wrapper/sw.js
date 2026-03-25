const CACHE_NAME = 'pp5smart-v1';
const STATIC_ASSETS = [
  '/',
  '/index.html',
  '/manifest.json',
  '/icons/icon-192.png',
  '/icons/icon-512.png'
];

// Install: cache static assets
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(STATIC_ASSETS))
  );
  self.skipWaiting();
});

// Activate: clean old caches
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Fetch: network-first for navigation, cache-first for static
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);
  
  // Don't cache Google Apps Script requests
  if (url.hostname.includes('google.com') || url.hostname.includes('googleapis.com')) {
    return;
  }
  
  // Static assets: cache-first
  if (STATIC_ASSETS.some(asset => url.pathname === asset || url.pathname === asset.replace(/\/$/, ''))) {
    event.respondWith(
      caches.match(event.request).then(cached => cached || fetch(event.request))
    );
    return;
  }
  
  // Everything else: network-first
  event.respondWith(
    fetch(event.request).catch(() => caches.match(event.request))
  );
});
