
// SERVICE WORKER v3 — EventPay
// v3: Never caches config.js, forces old cache clear

const CACHE_NAME = "eventpay-v3";
const NEVER_CACHE = ["config.js", "script.google.com"];

function toggleTheme() {
  const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
  document.documentElement.setAttribute('data-theme', isDark ? 'light' : 'dark');
  document.getElementById('themeToggle').textContent = isDark ? '🌙 Dark' : '☀️ Light';
  localStorage.setItem('theme', isDark ? 'light' : 'dark');
}
// Apply saved theme on load
(function() {
  const saved = localStorage.getItem('theme') || 'light';
  document.documentElement.setAttribute('data-theme', saved);
  const btn = document.getElementById('themeToggle');
  if (btn) btn.textContent = saved === 'dark' ? '☀️ Light' : '🌙 Dark';
})();
const CACHE_FILES = [
  "./index.html",
  "./status.html",
  "./complaint.html",
  "./admin-login.html",
  "./admin.html",
  "./shared.css",
  "./manifest.json",
  "./offline.html",
  "./icons/icon-192.png",
  "./icons/icon-512.png"
];

// INSTALL — cache core app files (NOT config.js)
self.addEventListener("install", event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(CACHE_FILES))
      .then(() => self.skipWaiting())
      .catch(() => self.skipWaiting())
  );
});

// ACTIVATE — delete ALL old caches
self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// FETCH — smart routing
self.addEventListener("fetch", event => {
  if (event.request.method !== "GET") return;

  const url = event.request.url;

  // 1. Google Apps Script API — ALWAYS go to network, never cache
  if (url.includes("script.google.com")) {
    event.respondWith(
      fetch(event.request).catch(() =>
        new Response(JSON.stringify({ error: "offline" }), {
          headers: { "Content-Type": "application/json" }
        })
      )
    );
    return;
  }

  // 2. config.js — ALWAYS fetch fresh from network (never use cache)
  if (url.includes("config.js")) {
    event.respondWith(
      fetch(event.request).catch(() => caches.match("./config.js"))
    );
    return;
  }

  // 3. Everything else — cache first, then network
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        if (response && response.status === 200) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(c => c.put(event.request, clone));
        }
        return response;
      }).catch(() => caches.match("./offline.html"));
    })
  );
});

