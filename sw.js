const CACHE_NAME = 'wabx-keygen-cache-v2'; // Incremented version
const urlsToCache = [
  '.', // This caches the root (index.html)
  'index.html',
  'manifest.json',
  'bin/icon.png',
  'bin/background.png',
  'bin/load.gif',
  'bin/welcome.mp3',
  'bin/success.mp3',
  'https://cdn.tailwindcss.com'
];

// Install event
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('Opened cache');
        return cache.addAll(urlsToCache);
      })
  );
});

// Fetch event (network-first strategy)
self.addEventListener('fetch', event => {
  event.respondWith(
    fetch(event.request)
      .catch(() => {
        // If network fails, try to get it from the cache
        return caches.match(event.request);
      })
  );
});

// Clean up old caches
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.filter(cacheName => {
          return cacheName.startsWith('wabx-keygen-cache-') &&
                 cacheName !== CACHE_NAME;
        }).map(cacheName => {
          return caches.delete(cacheName);
        })
      );
    })
  );
});
