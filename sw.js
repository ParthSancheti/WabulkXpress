// A very basic service worker to make the app installable (PWA)

const CACHE_NAME = 'wabx-keygen-cache-v1';
const urlsToCache = [
  '/',
  '/index.html', // Or whatever your main HTML file is
  'https://cdn.tailwindcss.com'
  // You can add more assets here if you want, like 'bin/icon.png'
  // But for a simple app like this, a minimal cache is fine.
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