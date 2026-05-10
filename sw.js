const CACHE_STATIC = 'dentat-static-v1';
const CACHE_TILES  = 'dentat-tiles-v1';
const MAX_TILES    = 200;

const STATIC_ASSETS = [
  './dentat.html',
  './manifest.json',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js',
  'https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css',
  'https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css',
  'https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js',
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_STATIC)
      .then(c => c.addAll(STATIC_ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(k => k !== CACHE_STATIC && k !== CACHE_TILES)
            .map(k => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const url = e.request.url;

  // Không cache: Google Sheet CSV, Apps Script (luôn cần dữ liệu mới nhất)
  if (url.includes('docs.google.com') || url.includes('script.google.com')) {
    e.respondWith(fetch(e.request));
    return;
  }

  // Cache tile bản đồ (OSM + Google)
  if (url.includes('tile.openstreetmap.org') || url.includes('google.com/vt')) {
    e.respondWith(
      caches.open(CACHE_TILES).then(async cache => {
        const cached = await cache.match(e.request);
        if (cached) return cached;
        const response = await fetch(e.request);
        if (response.ok) {
          cache.put(e.request, response.clone());
          trimCache(cache, MAX_TILES);
        }
        return response;
      }).catch(() => new Response('', { status: 503 }))
    );
    return;
  }

  // Network-first cho HTML (luôn lấy phiên bản mới nhất, fallback cache khi offline)
  if (url.endsWith('dentat.html') || url.endsWith('/Den%20tat/') || url.endsWith('/Den%20tat')) {
    e.respondWith(
      fetch(e.request)
        .then(res => {
          const clone = res.clone();
          caches.open(CACHE_STATIC).then(c => c.put(e.request, clone));
          return res;
        })
        .catch(() => caches.match(e.request))
    );
    return;
  }

  // Cache-first cho static assets còn lại (CSS, JS, fonts)
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request))
  );
});

async function trimCache(cache, maxItems) {
  const keys = await cache.keys();
  if (keys.length > maxItems) {
    await cache.delete(keys[0]);
  }
}
