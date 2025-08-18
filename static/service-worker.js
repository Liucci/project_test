// static/service-worker.js
self.addEventListener('install', (event) => {
  console.log('Service Worker installing.');
});

self.addEventListener('fetch', (event) => {
  // ここでキャッシュ戦略を指定可能
});
