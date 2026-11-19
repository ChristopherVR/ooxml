/*
 * Only generated app resources are cached. Documents remain in profile IndexedDB.
 *
 * Every page of this origin can write Cache Storage, so nothing read back from it is trusted until
 * this script has checked it. FILES maps each file of this build to its SHA-256, and a saved copy
 * is served only while it matches. A window still running an older build after an update keeps
 * loading that build's files from its cache, checked the same way against that build's own table:
 * each build saves its table, which must hash to the VERSION in its cache's name, and marks every
 * window it serves with a hash of the window's client id, which pages never see, so a page cannot
 * pass off a table or a cache of its own as the build a window runs.
 */
const BASE = new URL('./', self.location.href);
const PREFIX = `ooxml-app-${BASE.pathname}-`;
const CACHE = PREFIX + VERSION;
/** Where a build's cache keeps its table, written last, so only a finished install has one. */
const TABLE = new URL('.ooxml-sw/files', BASE).href;
const MARKS = new URL('.ooxml-sw/client/', BASE).href;
const CHUNKS = new URL('suite-assets/', BASE).href;
const table = (files) =>
	new Map(Object.entries(files).map(([file, hash]) => [new URL(file, BASE).href, hash]));
const current = { cache: CACHE, version: VERSION, files: table(FILES) };
/** How long a page load waits for the network before it falls back to the saved shell. */
const NETWORK_WAIT = 4000;

const sha256 = async (data) =>
	new Uint8Array(
		await crypto.subtle.digest(
			'SHA-256',
			typeof data === 'string' ? new TextEncoder().encode(data) : data,
		),
	);
const hex = (bytes) => Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join('');

/** Whether `response` is the file `build` shipped at `href`. */
async function genuine(build, href, response) {
	const hash = build.files.get(href);
	if (!hash) return false;
	const bytes = await sha256(await response.clone().arrayBuffer());
	return btoa(String.fromCharCode(...bytes)) === hash;
}
/** Saves a network response when it is this build's file; anything else is not saved. */
async function keep(href, response) {
	if (response.ok && !response.redirected && (await genuine(current, href, response)))
		await (await caches.open(CACHE)).put(href, response.clone());
	return response;
}
/** `build`'s saved copy of `href`, or undefined when there is none or it has been changed. */
async function saved(build, href) {
	const cache = await caches.open(build.cache);
	const response = await cache.match(href);
	if (!response) return undefined;
	if (await genuine(build, href, response)) return response;
	await cache.delete(href);
	return undefined;
}

/** The address that marks `clientId` as a window (or worker) running `build`. */
const mark = async (build, clientId) => MARKS + hex(await sha256(`${clientId}\n${build.version}`));
const runs = async (build, clientId) =>
	!!(await (await caches.open(build.cache)).match(await mark(build, clientId)));
/** A mark holds when it was written, as `Date.now()` text. */
async function remember(build, clientId) {
	if (clientId)
		await (
			await caches.open(build.cache)
		).put(await mark(build, clientId), new Response(String(Date.now())));
}
/**
 * How long a mark outlives its client's absence from `clients.matchAll`. A window or worker is
 * marked when its page or script is served, before it is a client, so a prune that listed the
 * clients just before must not take a fresh mark for a closed client's.
 */
const MARK_GRACE = 60_000;
async function fresh(cache, href) {
	const age = Date.now() - Number(await (await cache.match(href))?.text());
	return age >= 0 && age < MARK_GRACE;
}
/** The other builds whose caches are still kept: older ones, or a newer one not yet active. */
async function others() {
	return (await caches.keys())
		.filter((name) => name.startsWith(PREFIX) && name !== CACHE)
		.map((name) => ({ cache: name, version: name.slice(PREFIX.length) }));
}
/** The other build marked as run by `clientId` that ships `href`, with its checked table. */
async function builtFor(clientId, href) {
	if (!clientId) return undefined;
	for (const build of await others()) {
		if (!(await runs(build, clientId))) continue;
		const text = await (await (await caches.open(build.cache)).match(TABLE))?.text();
		if (text === undefined || hex(await sha256(text)) !== build.version) continue;
		const files = table(JSON.parse(text));
		if (files.has(href)) return { ...build, files };
	}
	return undefined;
}
/**
 * Deletes the marks of clients that are gone, and each other build's cache once no window or
 * worker runs it. A mark younger than MARK_GRACE counts as a client that is still starting. Only
 * the activating worker knows every other build is older, so it alone also deletes a cache no
 * client was ever marked by, and one without a table (an unfinished install, or a build from
 * before tables were kept). Before that, such a cache may be a newer build waiting to activate.
 * `arriving` is the window a navigation is opening, which is not a client yet.
 * Resolves to the clients it checked.
 */
async function prune(activating, arriving) {
	const clients = await self.clients.matchAll({ includeUncontrolled: true, type: 'all' });
	const ids = clients.map((client) => client.id);
	if (arriving) ids.push(arriving);
	for (const build of [current, ...(await others())]) {
		const cache = await caches.open(build.cache);
		const live = new Set(await Promise.all(ids.map((id) => mark(build, id))));
		const marks = (await cache.keys())
			.map((request) => request.url)
			.filter((href) => href.startsWith(MARKS));
		const stale = [];
		for (const href of marks) if (!live.has(href) && !(await fresh(cache, href))) stale.push(href);
		const used = stale.length < marks.length;
		const finished = !!(await cache.match(TABLE));
		const older = activating ? !used || !finished : marks.length > 0 && !used;
		if (build !== current && older) {
			await caches.delete(build.cache);
			continue;
		}
		for (const href of stale) await cache.delete(href);
	}
	return clients;
}

self.addEventListener('install', (event) =>
	event.waitUntil(
		(async () => {
			const queue = [...current.files.keys()];
			await Promise.all(
				Array.from({ length: 4 }, async () => {
					while (queue.length) {
						const href = queue.shift();
						// A retry after an interrupted install keeps what it already saved.
						if (await saved(current, href)) continue;
						const response = await fetch(href, { cache: 'no-cache' });
						// A file that is missing, redirected or not this build's (a cache still serving
						// the previous deploy) fails the install: the working build stays in charge,
						// and the browser installs again at its next update check.
						if (!response.ok || response.redirected || !(await genuine(current, href, response))) {
							queue.length = 0;
							throw new Error(`App resource unavailable: ${href}`);
						}
						await (await caches.open(CACHE)).put(href, response);
					}
				}),
			);
			const cache = await caches.open(CACHE);
			const kept = new Set((await cache.keys()).map((request) => request.url));
			if (![...current.files.keys()].every((href) => kept.has(href)))
				throw new Error('App resources were removed during the install');
			await cache.put(TABLE, new Response(JSON.stringify(FILES)));
			// Pages load their shell from the network and every script and style URL is versioned,
			// so a new worker can take over at once. Open windows keep the code they loaded, and
			// their build's cache stays (see prune) for the chunks they have not loaded yet.
			await self.skipWaiting();
		})(),
	),
);
self.addEventListener('activate', (event) =>
	event.waitUntil(
		(async () => {
			const clients = await prune(true);
			// A window no build has marked was loaded before any worker ran (a first visit), so it
			// is taken as this build's. A wrong guess only offers files a build shipped, checked
			// against that build's table.
			const builds = await others();
			for (const client of clients) {
				let marked = false;
				for (const build of builds) marked ||= await runs(build, client.id);
				if (!marked) await remember(current, client.id);
			}
			await self.clients.claim();
		})(),
	),
);
/** Pages: the network first, so a reload after a deploy gets it; the saved shell when offline. */
async function page(event, href) {
	const network = fetch(event.request.url, { cache: 'no-cache' }).then((response) => {
		if (!response.ok) throw new Error('Page unavailable');
		return response;
	});
	network.catch(() => {});
	const timeout = new Promise((resolve) => setTimeout(resolve, NETWORK_WAIT));
	let response;
	try {
		response = await Promise.race([network, timeout]);
	} catch {}
	// A shell that is not this build's (a newer deploy) starts a window this build does not run.
	if (response && !(await genuine(current, href, response))) return response;
	response ??= await saved(current, href);
	if (!response) return network;
	await remember(current, event.resultingClientId);
	return response;
}
/** A file `build` ships, from its checked copy, else the network. A worker it starts is marked. */
async function file(event, build, href) {
	const copy = build && (await saved(build, href));
	if (copy) {
		await remember(build, event.resultingClientId);
		return copy;
	}
	const response = await fetch(event.request);
	return build === current ? keep(href, response) : response;
}
self.addEventListener('fetch', (event) => {
	if (event.request.method !== 'GET') return;
	const url = new URL(event.request.url);
	url.hash = '';
	if (url.pathname.endsWith('/')) url.pathname += 'index.html';
	if (event.request.mode === 'navigate') {
		url.search = '';
		if (!current.files.has(url.href)) return;
		event.respondWith(page(event, url.href));
		event.waitUntil(prune(false, event.resultingClientId));
		return;
	}
	if (current.files.has(url.href)) event.respondWith(file(event, current, url.href));
	// A chunk this build does not ship is one a window of an older build asks for after an update.
	else if (url.href.startsWith(CHUNKS))
		event.respondWith(
			(async () => file(event, await builtFor(event.clientId, url.href), url.href))(),
		);
});
