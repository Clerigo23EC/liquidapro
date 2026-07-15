const CACHE_NAME = 'liquida-pro-v1';
const ASSETS = [
  'index.html',
  'manifest.json'
];

// Instalar y guardar los archivos en el teléfono
self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS);
    })
  );
});

// Activar
self.addEventListener('activate', (e) => {
  e.waitUntil(self.clients.claim());
});

// Responder desde el caché del teléfono si no hay internet (Anti-Dinosaurio)
self.addEventListener('fetch', (e) => {
  e.respondWith(
    caches.match(e.request).then((cachedResponse) => {
      if (cachedResponse) {
        return cachedResponse;
      }
      return fetch(e.request);
    }).catch(() => {
      // Retorna el archivo index local guardado si todo falla
      return caches.match('index.html');
    })
  );
});