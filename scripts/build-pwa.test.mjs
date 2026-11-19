import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { after, before, test } from 'node:test';
import vm from 'node:vm';
import { buildPwa } from './build-pwa.mjs';

const SHELL = `<!doctype html><html><head><meta charset="utf-8" />
<meta name="ooxml-root" content="./" />
<script src="appearance-init.js"></script>
<link rel="stylesheet" href="styles.css" />
<link rel="stylesheet" href="workspace.css" />
<script type="module" src="suite.js"></script>
</head><body></body></html>`;
const ORIGIN = 'https://pages.test/ooxml/';
/** Two deploys: B renames the boot chunk, as a later build whose chunk names changed does. */
const builds = {};

async function deploy(boot) {
	const out = mkdtempSync(join(tmpdir(), 'pwa-'));
	mkdirSync(join(out, 'suite-assets'));
	const files = {
		'index.html': SHELL,
		'appearance-init.js': 'document.documentElement.dataset.theme = "light";',
		'styles.css': 'body { margin: 0 }',
		'workspace.css': "@import url('./styles.css');",
		'suite.js': `import './suite-assets/${boot}';\n`,
		[`suite-assets/${boot}`]: `globalThis.booted = ${JSON.stringify(boot)};`,
	};
	for (const [name, text] of Object.entries(files)) writeFileSync(join(out, name), text);
	await buildPwa(out);
	return out;
}
before(async () => {
	builds.A = await deploy('boot-1.js');
	builds.B = await deploy('boot-2.js');
});
after(() => {
	for (const out of Object.values(builds)) rmSync(out, { recursive: true, force: true });
});

const read = (out, name) => readFileSync(join(out, name));
const sri = (out, name) =>
	`sha384-${createHash('sha384').update(read(out, name)).digest('base64')}`;
const path = (href) => new URL(href).pathname.slice(new URL(ORIGIN).pathname.length);
/** The site as the host serves `out`: 404 for a missing file. */
const site = (out) => (href) =>
	existsSync(join(out, path(href)))
		? new Response(read(out, path(href)))
		: new Response('', { status: 404 });
const offline = () => {
	throw new TypeError('Failed to fetch');
};

test('every shell script and style tag carries the integrity of the file it loads', () => {
	for (const page of ['index.html', 'apps/word/index.html']) {
		const html = read(builds.A, page).toString();
		for (const name of ['appearance-init.js', 'styles.css', 'workspace.css', 'suite.js']) {
			const tag = html.match(new RegExp(`<[^>]*${name.replace('.', '\\.')}\\?v=[^>]*>`, 'u'))?.[0];
			assert.ok(tag, `${page} loads ${name}`);
			assert.ok(tag.includes(`integrity="${sri(builds.A, name)}"`), `${page}: ${tag}`);
		}
	}
});

/** The workers' clock; a test moves it forward to let time pass. */
const clock = { now: Date.now() };
/** How long the worker keeps a mark whose client it does not list (MARK_GRACE). */
const MARK_GRACE = 60_000;

/**
 * A deploy's sw.js run against an in-memory Cache Storage, network and set of open windows, as a
 * browser would. `stores` is Cache Storage, shared by every worker of the origin.
 */
function worker(out, { network = site(out), stores = new Map(), clients = [] } = {}) {
	const open = async (name) => {
		if (!stores.has(name)) stores.set(name, new Map());
		const store = stores.get(name);
		return {
			match: async (href) => store.get(href)?.clone(),
			put: async (href, response) => void store.set(href, response),
			delete: async (href) => store.delete(href),
			keys: async () => [...store.keys()].map((url) => ({ url })),
		};
	};
	const listeners = {};
	const self = {
		location: { href: `${ORIGIN}sw.js` },
		addEventListener: (type, listener) => (listeners[type] = listener),
		skipWaiting: async () => {},
		clients: { claim: async () => {}, matchAll: async () => clients.map((id) => ({ id })) },
	};
	const fetches = [];
	const context = {
		self,
		caches: {
			open,
			keys: async () => [...stores.keys()],
			delete: async (name) => stores.delete(name),
		},
		fetch: async (input) => {
			const href = typeof input === 'string' ? input : input.url;
			fetches.push(href);
			return network(href);
		},
		crypto: globalThis.crypto,
		btoa,
		URL,
		Map,
		Set,
		Promise,
		Response,
		TextEncoder,
		JSON,
		Date: { now: () => clock.now },
		setTimeout,
		Uint8Array,
		Error,
		String,
		Object,
		Array,
	};
	vm.runInNewContext(read(out, 'sw.js').toString(), context);
	const run = async (type, event = {}) => {
		const waited = [];
		let responded;
		listeners[type]({
			...event,
			waitUntil: (promise) => waited.push(promise),
			respondWith: (promise) => (responded = promise),
		});
		await Promise.all(waited);
		return responded && (await responded);
	};
	const name = `ooxml-app-/ooxml/-${
		read(out, 'sw.js')
			.toString()
			.match(/VERSION="(\w+)"/u)[1]
	}`;
	return { run, stores, fetches, name, cache: () => stores.get(name) };
}

const request = (url) => ({ method: 'GET', mode: 'no-cors', url });
const navigate = (url, resultingClientId = '') => ({
	request: { method: 'GET', mode: 'navigate', url },
	resultingClientId,
});
const chunk = (name, clientId) => ({ request: request(`${ORIGIN}suite-assets/${name}`), clientId });
const appKeys = (stores) => [...stores.keys()].filter((key) => key.startsWith('ooxml-app-'));

/** Build A installed, active and serving one open window, `w1`. */
async function runningA(clients = ['w1']) {
	const a = worker(builds.A, { clients });
	await a.run('install');
	await a.run('activate');
	await a.run('fetch', navigate(`${ORIGIN}index.html?suite`, 'w1'));
	return a;
}

test('the worker serves a saved file only while it is the one the build shipped', async () => {
	const { run, fetches, cache } = await runningA();
	const store = cache();
	const suite = [...store.keys()].find((href) => /suite\.js\?v=/u.test(href));
	assert.ok(suite, 'suite.js is saved at install');

	fetches.length = 0;
	const genuine = await run('fetch', { request: request(suite) });
	assert.equal(await genuine.text(), read(builds.A, 'suite.js').toString());
	assert.deepEqual(fetches, [], 'a genuine copy comes from the cache');

	store.set(suite, new Response('globalThis.changed = true;'));
	const served = await run('fetch', { request: request(suite) });
	assert.equal(await served.text(), read(builds.A, 'suite.js').toString());
	assert.deepEqual(fetches, [suite], 'a changed copy is fetched again');
	assert.equal(
		await store.get(suite).text(),
		read(builds.A, 'suite.js').toString(),
		'and replaced',
	);
});

test('a changed saved shell is not served when the network is down', async () => {
	const online = await runningA();
	const page = `${ORIGIN}index.html`;
	online.cache().set(page, new Response('<meta name="ooxml-root"><script>changed()</script>'));
	const down = worker(builds.A, { network: offline, stores: online.stores });
	const served = await down.run('fetch', navigate(page)).catch((error) => error);
	assert.ok(served instanceof TypeError, 'the browser shows its own offline page');
	assert.equal(online.cache().has(page), false, 'the changed shell is deleted');

	await online.run('install');
	const restored = await worker(builds.A, { network: offline, stores: online.stores }).run(
		'fetch',
		navigate(page),
	);
	assert.equal(await restored.text(), read(builds.A, 'index.html').toString(), 'a genuine shell');
});

test('a cache without a file table is dropped, since its files cannot be checked', async () => {
	const { run, stores } = worker(builds.A);
	stores.set(`ooxml-app-/ooxml/-older`, new Map([[`${ORIGIN}suite.js?v=old`, new Response('x')]]));
	await run('install');
	await run('activate');
	assert.equal(appKeys(stores).length, 1);
});

test('a window still on the previous build keeps loading its chunks after an update', async () => {
	const a = await runningA();
	const b = worker(builds.B, { stores: a.stores, clients: ['w1'] });
	await b.run('install');
	await b.run('activate');
	assert.equal(appKeys(a.stores).length, 2, 'the previous build stays while a window runs it');

	const old = await b.run('fetch', chunk('boot-1.js', 'w1'));
	assert.equal(await old.text(), read(builds.A, 'suite-assets/boot-1.js').toString());
	assert.deepEqual(
		b.fetches.filter((href) => href.includes('boot-1')),
		[],
		'from the cache',
	);
	const offlineB = worker(builds.B, { network: offline, stores: a.stores, clients: ['w1'] });
	const offlineOld = await offlineB.run('fetch', chunk('boot-1.js', 'w1'));
	assert.equal(await offlineOld.text(), read(builds.A, 'suite-assets/boot-1.js').toString());

	const stranger = await b.run('fetch', chunk('boot-1.js', 'w2'));
	assert.equal(stranger.status, 404, 'only a window the previous build served gets its files');

	// Once the window has reloaded or closed, a page load after the grace deletes the previous build.
	clock.now += MARK_GRACE;
	const later = worker(builds.B, { stores: a.stores, clients: ['w3'] });
	await later.run('fetch', navigate(`${ORIGIN}index.html`, 'w3'));
	assert.deepEqual(appKeys(a.stores), [b.name]);
});

test("an older build's files and table are checked against the version that listed them", async () => {
	const evil = 'globalThis.owned = true;';
	const digest = (text, encoding) => createHash('sha256').update(text).digest(encoding);

	// A changed copy of the old chunk is not served, and is deleted.
	let a = await runningA();
	let b = worker(builds.B, { stores: a.stores, clients: ['w1'] });
	await b.run('install');
	await b.run('activate');
	const boot = `${ORIGIN}suite-assets/boot-1.js`;
	a.cache().set(boot, new Response(evil));
	assert.equal((await b.run('fetch', chunk('boot-1.js', 'w1'))).status, 404);
	assert.equal(a.cache().has(boot), false);

	// A table rewritten to accept a changed copy no longer hashes to the build's version.
	a = await runningA();
	b = worker(builds.B, { stores: a.stores, clients: ['w1'] });
	await b.run('install');
	await b.run('activate');
	const tableHref = `${ORIGIN}.ooxml-sw/files`;
	const files = JSON.parse(await a.cache().get(tableHref).clone().text());
	files['suite-assets/boot-1.js'] = digest(evil, 'base64');
	a.cache().set(tableHref, new Response(JSON.stringify(files)));
	a.cache().set(boot, new Response(evil));
	assert.equal((await b.run('fetch', chunk('boot-1.js', 'w1'))).status, 404);

	// A whole cache a page made up, with a table that matches its name, marks no real window.
	const forged = JSON.stringify(files);
	const name = `ooxml-app-/ooxml/-${digest(forged, 'hex')}`;
	a.stores.set(
		name,
		new Map([
			[tableHref, new Response(forged)],
			[boot, new Response(evil)],
		]),
	);
	assert.equal((await b.run('fetch', chunk('boot-1.js', 'w1'))).status, 404);
});

test('an install with a missing or stale file fails and the working build stays offline', async () => {
	let a;
	const index = `${ORIGIN}index.html`;
	const failures = {
		missing: (href) => (href.includes('styles.css') ? new Response('', { status: 500 }) : null),
		// An edge still serving the previous deploy's copy of a file.
		stale: (href) => (href === index ? site(builds.A)(href) : null),
	};
	for (const [kind, failure] of Object.entries(failures)) {
		a = await runningA();
		const b = worker(builds.B, {
			stores: a.stores,
			clients: ['w1'],
			network: (href) => failure(href) ?? site(builds.B)(href),
		});
		await assert.rejects(b.run('install'), /App resource unavailable/u, kind);
		const down = worker(builds.A, { network: offline, stores: a.stores, clients: ['w1'] });
		const shell = await down.run('fetch', navigate(index, 'w2'));
		assert.equal(await shell.text(), read(builds.A, 'index.html').toString(), `${kind}: A offline`);
	}

	// The next update check installs the build, reusing the files the failed attempts saved.
	const b = worker(builds.B, { stores: a.stores, clients: ['w1'] });
	await b.run('install');
	assert.ok(!b.fetches.some((href) => href.includes('appearance-init')), 'saved files are kept');
	await b.run('activate');
	const down = worker(builds.B, { network: offline, stores: a.stores, clients: ['w1'] });
	const shell = await down.run('fetch', navigate(index, 'w2'));
	assert.equal(await shell.text(), read(builds.B, 'index.html').toString(), 'B offline');
	const boot = await down.run('fetch', chunk('boot-2.js', 'w2'));
	assert.equal(await boot.text(), read(builds.B, 'suite-assets/boot-2.js').toString());
});

test('the worker version is the hash of the file table it saves', () => {
	const sw = read(builds.A, 'sw.js').toString();
	const version = sw.match(/VERSION="(\w+)"/u)[1];
	const table = sw.match(/const FILES=(.*);\n/u)[1];
	assert.equal(
		createHash('sha256')
			.update(JSON.stringify(JSON.parse(table)))
			.digest('hex'),
		version,
	);
});

test('a page load the previous build handles before the update activates keeps the new build', async () => {
	const a = await runningA();
	const b = worker(builds.B, { stores: a.stores, clients: ['w1'] });
	await b.run('install');
	// The previous build still controls the page until the new one activates.
	const before = worker(builds.A, { stores: a.stores, clients: ['w1', 'w4'] });
	await before.run('fetch', navigate(`${ORIGIN}index.html`, 'w4'));
	assert.ok(appKeys(a.stores).includes(b.name), 'the installed build is kept');
	await b.run('activate');

	const down = worker(builds.B, { network: offline, stores: a.stores, clients: ['w5'] });
	const shell = await down.run('fetch', navigate(`${ORIGIN}index.html`, 'w5'));
	assert.equal(await shell.text(), read(builds.B, 'index.html').toString(), 'B works offline');
	const boot = await down.run('fetch', chunk('boot-2.js', 'w5'));
	assert.equal(await boot.text(), read(builds.B, 'suite-assets/boot-2.js').toString());
});

const marksOf = (cache) => [...cache.keys()].filter((href) => href.includes('/.ooxml-sw/client/'));

test('page loads keep one mark per open window', async () => {
	const a = await runningA();
	for (let i = 0; i < 20; i++) {
		// Each load opens a window that is not a client yet, and the one before it has closed.
		clock.now += MARK_GRACE;
		const load = worker(builds.A, { stores: a.stores, clients: ['w1'] });
		await load.run('fetch', navigate(`${ORIGIN}index.html`, `n${i}`));
	}
	assert.equal(marksOf(a.cache()).length, 2, 'w1 and the window the last load opened');

	// The window the last load opened still gets its build's chunks after an update.
	const b = worker(builds.B, { stores: a.stores, clients: ['w1', 'n19'] });
	await b.run('install');
	await b.run('activate');
	const old = await b.run('fetch', chunk('boot-1.js', 'n19'));
	assert.equal(await old.text(), read(builds.A, 'suite-assets/boot-1.js').toString());
});

test('a build no window runs is deleted at the next activation, over several deploys', async () => {
	const a = await runningA();
	const b = worker(builds.B, { stores: a.stores, clients: ['w1'] });
	await b.run('install');
	await b.run('activate');
	// A rollback to A with only a window that B never served open.
	const back = worker(builds.A, { stores: a.stores, clients: ['w9'] });
	await back.run('install');
	await back.run('activate');
	assert.deepEqual(appKeys(a.stores), [a.name]);
});

const INDEX = `${ORIGIN}index.html`;

// Two windows open at once (a restored session): w3's navigation lists the clients before w2 is
// one, after w2's page response marked it.
test('a window opened during another navigation keeps its build after an update', async () => {
	const a = await runningA();
	await worker(builds.A, { stores: a.stores, clients: ['w1'] }).run('fetch', navigate(INDEX, 'w2'));
	await worker(builds.A, { stores: a.stores, clients: ['w1'] }).run('fetch', navigate(INDEX, 'w3'));
	const b = worker(builds.B, { stores: a.stores, clients: ['w1', 'w2', 'w3'] });
	await b.run('install');
	await b.run('activate');
	const old = await b.run('fetch', chunk('boot-1.js', 'w2'));
	assert.equal(old.status, 200, `w2 gets A's chunk, got ${old.status}`);
	assert.equal(await old.text(), read(builds.A, 'suite-assets/boot-1.js').toString());
});

// A dedicated worker is marked when its script is served, before it is a client.
test("a worker started during a navigation's prune keeps its build after an update", async () => {
	const a = await runningA();
	const script = [...a.cache().keys()].find((href) => /suite\.js\?v=/u.test(href));
	await a.run('fetch', { request: request(script), resultingClientId: 'k1' });
	await worker(builds.A, { stores: a.stores, clients: ['w1'] }).run('fetch', navigate(INDEX, 'w3'));
	const b = worker(builds.B, { stores: a.stores, clients: ['w1', 'k1', 'w3'] });
	await b.run('install');
	await b.run('activate');
	const old = await b.run('fetch', chunk('boot-1.js', 'k1'));
	assert.equal(old.status, 200, `k1 gets A's chunk, got ${old.status}`);
});

test('a starting window keeps an older build until its mark goes stale', async () => {
	const a = await runningA();
	clock.now += MARK_GRACE;
	await worker(builds.A, { stores: a.stores, clients: ['w1'] }).run('fetch', navigate(INDEX, 'w2'));
	const b = worker(builds.B, { stores: a.stores, clients: [] });
	await b.run('install');
	// w1 has closed and w2 is not listed yet: its fresh mark keeps A.
	await b.run('activate');
	assert.ok(appKeys(a.stores).includes(a.name), "A's cache is kept for w2");
	clock.now += MARK_GRACE;
	await worker(builds.B, { stores: a.stores, clients: [] }).run('fetch', navigate(INDEX, 'w4'));
	assert.deepEqual(appKeys(a.stores), [b.name], 'A goes once no client runs it');
});
