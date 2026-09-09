/**
 * PrimaNota Cassa Dolce Vita — Service Worker
 *
 * Volutamente SENZA cache: l'app ha comunque bisogno della rete per leggere
 * e scrivere su Google Sheets, e una cache locale causerebbe il problema
 * delle versioni vecchie che restano in giro dopo un aggiornamento.
 * Serve solo a rendere l'app installabile sul telefono/PC.
 */

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  // sempre dalla rete, nessuna copia locale
  event.respondWith(fetch(event.request));
});
