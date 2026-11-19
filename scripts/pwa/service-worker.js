/*
 * Only generated app resources are cached. Documents remain in profile IndexedDB.
 *
 * Every page of this origin can write Cache Storage, so a saved copy is served only while it is
 * byte for byte the file this build shipped: FILES maps each one to its SHA-256, and a copy that
 * no longer matches is deleted and fetched again from the network.
 */
const BASE = new URL('./', self.location.href);
const PREFIX = `ooxml-app-${BASE.pathname}-`;
const CACHE = PREFIX + VERSION;
const hashes = new Map(
	Object.entries(FILES).map(([file, hash]) => [new URL(file, BASE).href, hash]),
);
/** How long a page load waits for the network before it falls back to the saved shell. */
const NETWORK_WAIT = 4000;

async function genuine(href, response) {
	const bytes = new Uint8Array(
		await crypto.subtle.digest('SHA-256', await response.clone().arrayBuffer()),
	);
	return btoa(String.fromCharCode(...bytes)) === hashes.get(href);
}
/** Saves a network response when it is this build's file; anything else is not saved. */
async function keep(href, response) {
	if (response.ok && !response.redirected && (await genuine(href, response)))
		await (await caches.open(CACHE)).put(href, response.clone());
	return response;
}
/** The saved copy of `href`, or undefined when there is none or it has been changed. */
async function saved(href) {
	const cache = await caches.open(CACHE);
	const response = await cache.match(href);
	if (!response) return undefined;
	if (await genuine(href, response)) return response;
	await cache.delete(href);
	return undefined;
}

self.addEventListener('install', (event) =>
	event.waitUntil(
		(async () => {
			const queue = [...hashes.keys()];
			await Promise.all(
				Array.from({ length: 4 }, async () => {
					while (queue.length) {
						const href = queue.shift();
						const response = await fetch(href, { cache: 'no-cache' });
						if (!response.ok || response.redirected) throw new Error('App resource unavailable');
						await keep(href, response);
					}
				}),
			);
			// Pages load their shell from the network and every script and style URL is versioned,
			// so a new worker can take over at once. Open windows keep the code they loaded.
			await self.skipWaiting();
		})(),
	),
);
self.addEventListener('activate', (event) =>
	event.waitUntil(
		(async () => {
			// An older build's files cannot be checked against this build's hashes, so they go.
			for (const key of await caches.keys())
				if (key.startsWith(PREFIX) && key !== CACHE) await caches.delete(key);
			await self.clients.claim();
		})(),
	),
);
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
		if (hashes.has(url.href)) event.respondWith(page(event.request, url.href));
		return;
	}
	if (!hashes.has(url.href)) return;
	event.respondWith(
		(async () => (await saved(url.href)) ?? keep(url.href, await fetch(event.request)))(),
	);
});
