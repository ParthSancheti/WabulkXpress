// Change the cache name every time you update any file
const CACHE_NAME = 'wabx-keygen-cache-v2';

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

// Install event: Opens the cache and adds all specified files
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('Opened cache');
        return cache.addAll(urlsToCache);
      })
  );
});

// Fetch event: Tries network first, then falls back to cache
self.addEventListener('fetch', event => {
  event.respondWith(
    fetch(event.request)
      .catch(() => {
        // If network fails (offline), try to get it from the cache
        return caches.match(event.request);
      })
  );
});

// Activate event: Cleans up old, unused caches
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.filter(cacheName => {
          // Find all caches that start with our name but are NOT the current one
          return cacheName.startsWith('wabx-keygen-cache-') &&
                 cacheName !== CACHE_NAME;
        }).map(cacheName => {
          // Delete the old cache
          console.log('Deleting old cache:', cacheName);
          return caches.delete(cacheName);
        })
      );
    })
  );
});
