// sw.js
const CACHE_NAME = 'wabx-keygen-cache-v1';
// Add the name of your HTML file if it's not 'index.html'
const ASSETS_TO_CACHE = [
  '.', // This caches the root URL (your HTML file)
  'bin/Logo.png',
  'https://images.unsplash.com/photo-1557683316-973673baf926?ixlib=rb-4.0.3&auto=format&fit=crop&w=1920&q=80'
];

// Install event: cache local assets
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('Cache opened');
        // Use addAll for non-CORS requests
        return cache.addAll(['.', 'bin/Logo.png']).then(() => {
          // Use add for CORS requests (like the background image)
          // 'no-cors' mode means we store it but can't inspect it
          return cache.add(new Request(ASSETS_TO_CACHE[2], { mode: 'no-cors' }));
        });
      })
  );
});

// Activate event: clean up old caches
self.addEventListener('activate', event => {
  const cacheWhitelist = [CACHE_NAME];
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cacheName => {
          if (cacheWhitelist.indexOf(cacheName) === -1) {
            return caches.delete(cacheName);
          }
        })
      );
    })
  );
});

// Fetch event: Network-first, then cache fallback
// This ensures API calls to Google Apps Script always try the network first.
self.addEventListener('fetch', event => {
  const requestUrl = new URL(event.request.url);

  // Always go to network for Google Apps Script
  if (requestUrl.hostname === 'script.google.com') {
    event.respondWith(fetch(event.request));
    return;
  }

  // For other requests, try network first, then fall back to cache
  event.respondWith(
    fetch(event.request)
      .then(networkResponse => {
        return networkResponse;
      })
      .catch(() => {
        // If the network fails, try to get it from the cache
        return caches.match(event.request);
      })
  );
});
