/* Only generated app resources are cached. Documents remain in profile IndexedDB. */
const BASE = new URL('./', self.location.href);
const CACHE = `ooxml-app-${BASE.pathname}-${VERSION}`;
const urls = new Set(FILES.map((file) => new URL(file, BASE).href));
async function fetchStatic(url) {
	const response = await fetch(url, { cache: 'no-cache' });
	if (!response.ok || response.redirected) throw new Error('App resource unavailable');
	if (url.endsWith('.html') && !(await response.clone().text()).includes('name="ooxml-root"'))
		throw new Error('Not the app shell');
	return response;
}
self.addEventListener('install', (event) =>
	event.waitUntil(
		(async () => {
			const cache = await caches.open(CACHE);
			const queue = [...urls];
			await Promise.all(
				Array.from({ length: 4 }, async () => {
					while (queue.length) {
						const url = queue.shift();
						await cache.put(url, await fetchStatic(url));
					}
				}),
			);
			// Updates activate after the previous app windows close, preserving active editing sessions.
		})(),
	),
);
self.addEventListener('activate', (event) =>
	event.waitUntil(
		(async () => {
			for (const key of await caches.keys())
				if (key.startsWith(`ooxml-app-${BASE.pathname}-`) && key !== CACHE)
					await caches.delete(key);
			await self.clients.claim();
		})(),
	),
);
self.addEventListener('fetch', (event) => {
	if (event.request.method !== 'GET') return;
	const url = new URL(event.request.url);
	url.hash = '';
	if (url.pathname.endsWith('/')) url.pathname += 'index.html';
	if (!urls.has(url.href)) return;
	event.respondWith(
		(async () => {
			const cache = await caches.open(CACHE);
			// Versioned shell and assets stay together until the next worker activates.
			const saved = await cache.match(url.href);
			return saved || fetch(event.request);
		})(),
	);
});
