// Guérim CRM : service worker minimal.
// Il ne met rien en cache : aucune donnée ne reste stockée sur l'appareil.
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', e => {
  // Seules les pages de l'appli passent par ici ; les échanges avec Apps Script vont directement au réseau.
  if (new URL(e.request.url).origin === self.location.origin) e.respondWith(fetch(e.request));
});
