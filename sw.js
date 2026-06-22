const CACHE = 'madryn-norte-v3';
const OFFLINE_HTML = `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Sin conexión — Madryn Norte</title><style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:system-ui,sans-serif;background:linear-gradient(135deg,#7f1d1d,#166534);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px}.card{background:#fff;border-radius:16px;padding:36px 28px;max-width:340px;text-align:center;box-shadow:0 8px 32px rgba(0,0,0,.2)}.ico{font-size:52px;margin-bottom:16px}.t{font-size:20px;font-weight:700;color:#0f2318;margin-bottom:10px}.s{font-size:14px;color:#4d6557;line-height:1.6}</style></head><body><div class="card"><div class="ico">📵</div><div class="t">Sin conexión</div><div class="s">Abrí Madryn Norte cuando tengas internet.</div></div></body></html>`;

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => {
      c.put('__offline__', new Response(OFFLINE_HTML, { headers: { 'Content-Type': 'text/html' } }));
      return c.addAll(['./escudo.jpg', './manifest.json']);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = e.request.url;
  // Nunca interceptar llamadas al backend de Google
  if (url.includes('script.google.com') || url.includes('googleusercontent.com')) return;

  const req = e.request;
  if (req.mode === 'navigate' || req.headers.get('accept')?.includes('text/html')) {
    // HTML: red primero, caché de respaldo, offline como último recurso
    e.respondWith(
      fetch(req)
        .then(r => { const c = r.clone(); caches.open(CACHE).then(ch => ch.put(req, c)); return r; })
        .catch(() => caches.match(req).then(r => r || caches.match('__offline__')))
    );
  } else {
    // Estáticos: caché primero
    e.respondWith(caches.match(req).then(r => r || fetch(req).catch(() => {})));
  }
});
