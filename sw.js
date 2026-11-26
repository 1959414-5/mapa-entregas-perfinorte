const CACHE_PREFIX = 'perfinorte-entregas-';
const CACHE_NAME = CACHE_PREFIX + 'v7';
const APP_BASE = new URL('./', self.registration.scope);

const APP_FILES = [
  'index.html',
  'vendedores.html',
  'motorista.html',
  'manifest.webmanifest',
  'favicon.svg',
  'pwa-register.js',
  'icons/icon-192.png',
  'icons/icon.svg',
  'icons/icon-maskable.svg',
  'icons/apple-touch-icon.png'
];

const PRECACHE_URLS = APP_FILES.map((file) => new URL(file, APP_BASE).href);
const STATIC_PATHS = new Set(PRECACHE_URLS.map((url) => new URL(url).pathname));

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(PRECACHE_URLS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(
        keys
          .filter((key) => key.startsWith(CACHE_PREFIX) && key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const request = event.request;
  if (request.method !== 'GET') return;

  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;
  if (!url.pathname.startsWith(APP_BASE.pathname)) return;

  // As páginas HTML usam network-first para receber atualizações quando online.
  if (request.mode === 'navigate') {
    event.respondWith((async () => {
      const cache = await caches.open(CACHE_NAME);
      try {
        const response = await fetch(request);
        if (response && response.ok) {
          await cache.put(request, response.clone());
          return response;
        }
      } catch (_) {
        // Sem rede: tenta entregar a página que foi armazenada na instalação.
      }

      const exact = await cache.match(request, { ignoreSearch: true });
      if (exact) return exact;

      const fallback = await cache.match(new URL('index.html', APP_BASE).href);
      if (fallback) return fallback;

      return new Response(
        '<!doctype html><html lang="pt-BR"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Perfinorte</title><body style="font:16px sans-serif;padding:32px;background:#111925;color:#eef1f4"><h1>Perfinorte</h1><p>Não foi possível abrir a página offline. Conecte-se à internet e tente novamente.</p></body></html>',
        { status: 503, headers: { 'Content-Type': 'text/html; charset=utf-8' } }
      );
    })());
    return;
  }

  // Somente manifesto, registro PWA e ícones são armazenados. Chamadas de API,
  // autenticação, dados de entregas e requisições externas nunca entram no cache.
  if (STATIC_PATHS.has(url.pathname)) {
    event.respondWith((async () => {
      const cache = await caches.open(CACHE_NAME);
      const cached = await cache.match(request);
      if (cached) return cached;

      const response = await fetch(request);
      if (response && response.ok) await cache.put(request, response.clone());
      return response;
    })());
  }
});