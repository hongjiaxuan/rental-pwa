// ═══════════════════════════════════════════════════
// Service Worker — 套房管理系統 PWA
// 快取策略：
//   靜態資源（HTML / CDN JS / CSS）→ Cache First
//   GAS API 呼叫（script.google.com）→ Network Only（永不快取）
// ═══════════════════════════════════════════════════

const CACHE_NAME = 'rental-mgmt-v3';

// 預快取的靜態資源
const STATIC_ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon/icon-192.png',
  './icon/icon-512.png',
  'https://cdn.tailwindcss.com',
  'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js'
];

// ── 安裝：預先快取靜態資源 ─────────────────────────
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(STATIC_ASSETS))
      .then(() => self.skipWaiting())
  );
});

// ── 啟用：清除舊版快取 ─────────────────────────────
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

// ── 攔截請求 ──────────────────────────────────────
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // GAS API / Google 服務 → 直接走網路，不快取
  if (
    url.hostname === 'script.google.com' ||
    url.hostname === 'accounts.google.com' ||
    url.hostname.endsWith('.googleapis.com')
  ) {
    event.respondWith(fetch(event.request));
    return;
  }

  // POST 請求 → 走網路
  if (event.request.method !== 'GET') {
    event.respondWith(fetch(event.request));
    return;
  }

  // 靜態資源 → Cache First（快取優先，失敗再走網路）
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        // 只快取成功的同源或 CDN 資源
        if (
          response.status === 200 &&
          (url.hostname === self.location.hostname ||
           url.hostname.includes('cdnjs.cloudflare.com') ||
           url.hostname === 'cdn.tailwindcss.com')
        ) {
          const cloned = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, cloned));
        }
        return response;
      }).catch(() => {
        // 網路斷線時回傳快取的 index.html（SPA fallback）
        if (event.request.destination === 'document') {
          return caches.match('./index.html');
        }
      });
    })
  );
});
