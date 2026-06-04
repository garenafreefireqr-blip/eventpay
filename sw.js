// ============================================================
// SERVICE WORKER — EventPay PWA
// This file makes the app work offline and installable
// ============================================================

const CACHE_NAME = "eventpay-v2";
const OFFLINE_URL = "./offline.html";

// Files to cache for offline use
const CACHE_FILES = [
  "./index.html",
  "./status.html",
  "./complaint.html",
  "./admin-login.html",
  "./admin.html",
  "./shared.css",
  "./config.js",
  "./manifest.json",
  "./offline.html",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "https://fonts.googleapis.com/css2?family=Sora:wght@300;400;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap",
  "https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js",
  "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js",
  "https://cdn.jsdelivr.net/npm/chart.js"
];

// ---- INSTALL: cache all files ----
self.addEventListener("install", event => {
  console.log("[SW] Installing...");
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log("[SW] Caching app files");
        // Cache local files strictly, external ones softly
        const localFiles = CACHE_FILES.filter(f => f.startsWith("./"));
        const externalFiles = CACHE_FILES.filter(f => !f.startsWith("./"));
        return cache.addAll(localFiles)
          .then(() => {
            // Cache external files one by one, ignoring failures
            return Promise.allSettled(
              externalFiles.map(url =>
                fetch(url).then(r => cache.put(url, r)).catch(() => {})
              )
            );
          });
      })
      .then(() => self.skipWaiting())
  );
});

// ---- ACTIVATE: clean old caches ----
self.addEventListener("activate", event => {
  console.log("[SW] Activating...");
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(key => key !== CACHE_NAME)
            .map(key => {
              console.log("[SW] Deleting old cache:", key);
              return caches.delete(key);
            })
      )
    ).then(() => self.clients.claim())
  );
});

// ---- FETCH: serve from cache, fall back to network ----
self.addEventListener("fetch", event => {
  // Skip non-GET and chrome-extension requests
  if (event.request.method !== "GET") return;
  if (event.request.url.startsWith("chrome-extension")) return;

  // Google Apps Script API — always go to network (never cache API calls)
  if (event.request.url.includes("script.google.com")) {
    event.respondWith(
      fetch(event.request).catch(() => {
        return new Response(
          JSON.stringify({ error: "You are offline. Please connect to internet." }),
          { headers: { "Content-Type": "application/json" } }
        );
      })
    );
    return;
  }

  // For everything else: try cache first, then network
  event.respondWith(
    caches.match(event.request)
      .then(cached => {
        if (cached) return cached;
        return fetch(event.request)
          .then(response => {
            // Cache new successful responses
            if (response && response.status === 200) {
              const clone = response.clone();
              caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
            }
            return response;
          })
          .catch(() => {
            // Offline fallback for HTML pages
            if (event.request.headers.get("accept")?.includes("text/html")) {
              return caches.match(OFFLINE_URL);
            }
          });
      })
  );
});

// ---- PUSH NOTIFICATION support (future use) ----
self.addEventListener("push", event => {
  if (!event.data) return;
  const data = event.data.json();
  self.registration.showNotification(data.title || "EventPay", {
    body: data.body || "You have a new notification",
    icon: "./icons/icon-192.png",
    badge: "./icons/icon-96.png",
    vibrate: [200, 100, 200]
  });
});
