const CACHE='teatro-v131';
const ASSETS=['./sw.js','./manifest.json'];

self.addEventListener('install', e=>{
  e.waitUntil(caches.open(CACHE).then(c=>c.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener('activate', e=>{
  e.waitUntil(caches.keys().then(keys=>Promise.all(
    keys.filter(k=>k!==CACHE).map(k=>caches.delete(k))
  )));
  self.clients.claim();
});

self.addEventListener('fetch', e=>{
  const url = new URL(e.request.url);
  // index.html sempre dalla rete (network-first) — garantisce aggiornamenti immediati
  if(url.pathname.endsWith('/') || url.pathname.endsWith('index.html')){
    e.respondWith(
      fetch(e.request).catch(()=>caches.match('./index.html'))
    );
    return;
  }
  // Tutto il resto: cache-first
  e.respondWith(caches.match(e.request).then(cached=>cached||fetch(e.request)));
});

self.addEventListener('message', e=>{
  if(e.data==='skipWaiting') self.skipWaiting();
});
