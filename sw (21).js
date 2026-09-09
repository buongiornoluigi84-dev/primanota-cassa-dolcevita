/**
 * PrimaNota Cassa Dolce Vita — Service Worker
 *
 * Strategia "rete per prima": ogni richiesta va sempre online, così un
 * aggiornamento si vede subito senza problemi di versioni vecchie.
 * La copia locale entra in gioco SOLO se la rete non risponde — serve
 * perché Chrome installa l'app come applicazione vera solo se questa
 * riesce ad aprirsi anche offline.
 */
const CACHE = 'primanota-v1';
const FALLBACK = './';

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE).then(c => c.addAll([FALLBACK, 'index.html']).catch(() => {}))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;

  // Le chiamate a Google Apps Script non vanno mai messe in cache
  if (req.url.includes('script.google.com')) return;

  event.respondWith(
    fetch(req)
      .then(res => {
        if (res && res.ok && req.url.startsWith(self.location.origin)) {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
        }
        return res;
      })
      .catch(() =>
        caches.match(req).then(hit => hit || caches.match(FALLBACK))
      )
  );
});
