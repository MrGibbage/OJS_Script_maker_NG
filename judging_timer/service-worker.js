const CACHE_NAME = 'judging-timer-v1';
const ASSETS = [
  './index.html',
  './manifest.json',
  './icons/icon-192.svg',
  './icons/icon-512.svg'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.map(k => { if(k !== CACHE_NAME) return caches.delete(k); }))).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  // serve from cache first for app shell and icons
  if(ASSETS.some(p => url.pathname.endsWith(p.replace('./','')))){
    event.respondWith(caches.match(event.request).then(resp => resp || fetch(event.request)));
    return;
  }
  // for HTML navigation requests, return offline page if network fails
  if(event.request.mode === 'navigate' || (event.request.headers.get('accept') || '').includes('text/html')){
    event.respondWith(fetch(event.request).catch(() => caches.match('./index.html')));
    return;
  }
  // otherwise try network then cache
  event.respondWith(fetch(event.request).catch(() => caches.match(event.request)));
});
