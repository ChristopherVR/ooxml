/* Only generated app resources are cached. Documents remain in profile IndexedDB. */
const BASE = new URL('./', self.location.href);
const PREFIX = `ooxml-app-${BASE.pathname}-`;
const CACHE = PREFIX + VERSION;
const urls = new Set(FILES.map((file) => new URL(file, BASE).href));
const chunks = new URL('suite-assets/', BASE).href;
/** How long a page load waits for the network before it falls back to the saved shell. */
const NETWORK_WAIT = 4000;
async function fetchStatic(url) {
	const response = await fetch(url, { cache: 'no-cache' });
	if (!response.ok || response.redirected) throw new Error('App resource unavailable');
	if (url.includes('.html') && !(await response.clone().text()).includes('name="ooxml-root"'))
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
			// Pages load their shell from the network and every script and style URL is versioned,
			// so a new worker can take over at once. Open windows keep the code they loaded, and
			// the previous cache stays (see activate) for the chunks they have not loaded yet.
			await self.skipWaiting();
		})(),
	),
);
self.addEventListener('activate', (event) =>
	event.waitUntil(
		(async () => {
			// Keys come back in creation order: keep this version and the one before it.
			const older = (await caches.keys()).filter((key) => key.startsWith(PREFIX) && key !== CACHE);
			for (const key of older.slice(0, -1)) await caches.delete(key);
			await self.clients.claim();
		})(),
	),
);
/** A response saved by this version or the one before it (chunks an open window still needs). */
async function saved(href) {
	const current = await (await caches.open(CACHE)).match(href);
	if (current) return current;
	for (const key of await caches.keys())
		if (key.startsWith(PREFIX)) {
			const old = await (await caches.open(key)).match(href);
			if (old) return old;
		}
	return undefined;
}
/** Pages: the network first, so a reload after a deploy gets it; the saved shell when offline. */
async function page(request, href) {
	const network = fetch(request.url, { cache: 'no-cache' }).then((response) => {
		if (!response.ok) throw new Error('Page unavailable');
		return response;
	});
	network.catch(() => {});
	const timeout = new Promise((resolve) => setTimeout(resolve, NETWORK_WAIT));
	try {
		const response = await Promise.race([network, timeout]);
		if (response) return response;
	} catch {}
	return (await saved(href)) ?? network;
}
self.addEventListener('fetch', (event) => {
	if (event.request.method !== 'GET') return;
	const url = new URL(event.request.url);
	url.hash = '';
	if (url.pathname.endsWith('/')) url.pathname += 'index.html';
	if (event.request.mode === 'navigate') {
		url.search = '';
		if (urls.has(url.href)) event.respondWith(page(event.request, url.href));
		return;
	}
	if (!urls.has(url.href) && !url.href.startsWith(chunks)) return;
	event.respondWith((async () => (await saved(url.href)) ?? fetch(event.request))());
});
