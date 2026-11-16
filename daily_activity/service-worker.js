const CACHE_NAME = "ritme-shell-v2";
const APP_SHELL = [
  "./",
  "./index.html",
  "./styles.css?v=1",
  "./config.js?v=1",
  "./app.js?v=2",
  "./manifest.webmanifest",
  "./icon.svg",
];
const SHELL_PATHS = new Set(
  APP_SHELL.map((asset) => new URL(asset, self.registration.scope).pathname),
);

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches
      .open(CACHE_NAME)
      .then((cache) => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting()),
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((names) =>
        Promise.all(
          names
            .filter((name) => name.startsWith("ritme-shell-") && name !== CACHE_NAME)
            .map((name) => caches.delete(name)),
        ),
      )
      .then(() => self.clients.claim()),
  );
});

self.addEventListener("fetch", (event) => {
  const request = event.request;
  const url = new URL(request.url);
  if (
    request.method !== "GET" ||
    url.origin !== self.location.origin ||
    url.pathname.startsWith(new URL("../api/daily-activity", self.registration.scope).pathname)
  ) {
    return;
  }

  const isNavigation = request.mode === "navigate";
  if (!isNavigation && !SHELL_PATHS.has(url.pathname)) return;
  event.respondWith(
    (async () => {
      const cache = await caches.open(CACHE_NAME);
      try {
        const response = await fetch(request);
        if (response.ok) await cache.put(request, response.clone());
        return response;
      } catch (error) {
        const cached = await cache.match(request, { ignoreSearch: true });
        if (cached) return cached;
        if (isNavigation) {
          const page = await cache.match(
            new URL("./index.html", self.registration.scope).href,
          );
          if (page) return page;
        }
        throw error;
      }
    })(),
  );
});
