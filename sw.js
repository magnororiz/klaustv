/* Klaus TV — service worker do app (PWA).
 * Regra: sempre busca na internet primeiro (o site nunca fica velho).
 * Sem internet, abre a última cópia guardada das páginas e imagens do próprio site.
 * Nunca guarda os dados do jogo (servidor Google): os números sempre vêm ao vivo. */
const CACHE = 'klaustv-app-v1';
const BASICO = ['aovivo.html', 'manifest.webmanifest', 'logo_klaustv_tv.png', 'bola_fiba.png', 'app/icone-192.png', 'app/apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(BASICO)).catch(() => {}).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request; const url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin) return; // servidor Google, mapas etc.: direto na internet
  e.respondWith(
    fetch(req).then(res => {
      if (res && res.ok) { const cp = res.clone(); caches.open(CACHE).then(c => c.put(req, cp)).catch(() => {}); }
      return res;
    }).catch(() => caches.match(req, { ignoreSearch: true }).then(r => r || caches.match('aovivo.html')))
  );
});
