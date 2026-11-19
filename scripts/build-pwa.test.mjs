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
let out;

before(async () => {
	out = mkdtempSync(join(tmpdir(), 'pwa-'));
	mkdirSync(join(out, 'suite-assets'));
	const files = {
		'index.html': SHELL,
		'appearance-init.js': 'document.documentElement.dataset.theme = "light";',
		'styles.css': 'body { margin: 0 }',
		'workspace.css': "@import url('./styles.css');",
		'suite.js': "import './suite-assets/boot-1.js';\n",
		'suite-assets/boot-1.js': 'globalThis.booted = true;',
	};
	for (const [name, text] of Object.entries(files)) writeFileSync(join(out, name), text);
	await buildPwa(out);
});
after(() => rmSync(out, { recursive: true, force: true }));

const read = (name) => readFileSync(join(out, name));
const sri = (name) => `sha384-${createHash('sha384').update(read(name)).digest('base64')}`;

test('every shell script and style tag carries the integrity of the file it loads', () => {
	for (const page of ['index.html', 'apps/word/index.html']) {
		const html = read(page).toString();
		for (const name of ['appearance-init.js', 'styles.css', 'workspace.css', 'suite.js']) {
			const tag = html.match(new RegExp(`<[^>]*${name.replace('.', '\\.')}\\?v=[^>]*>`, 'u'))?.[0];
			assert.ok(tag, `${page} loads ${name}`);
			assert.ok(tag.includes(`integrity="${sri(name)}"`), `${page}: ${tag}`);
		}
	}
});

/** sw.js run against an in-memory Cache Storage and network, as a browser would. */
function worker({ offline = false, stores = new Map() } = {}) {
	const open = async (name) => {
		if (!stores.has(name)) stores.set(name, new Map());
		const store = stores.get(name);
		return {
			match: async (href) => store.get(href)?.clone(),
			put: async (href, response) => void store.set(href, response),
			delete: async (href) => store.delete(href),
			keys: async () => [...store.keys()],
		};
	};
	const listeners = {};
	const self = {
		location: { href: `${ORIGIN}sw.js` },
		addEventListener: (type, listener) => (listeners[type] = listener),
		skipWaiting: async () => {},
		clients: { claim: async () => {} },
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
			if (offline) throw new TypeError('Failed to fetch');
			const path = new URL(href).pathname.slice(new URL(ORIGIN).pathname.length);
			return existsSync(join(out, path))
				? new Response(read(path))
				: new Response('', { status: 404 });
		},
		crypto: globalThis.crypto,
		btoa,
		URL,
		Map,
		Promise,
		setTimeout,
		Uint8Array,
		Error,
		String,
		Object,
		Array,
	};
	vm.runInNewContext(read('sw.js').toString(), context);
	const run = async (type, event = {}) => {
		let waited;
		let responded;
		listeners[type]({
			...event,
			waitUntil: (promise) => (waited = promise),
			respondWith: (promise) => (responded = promise),
		});
		await waited;
		return responded && (await responded);
	};
	return { run, stores, fetches };
}

const request = (url) => ({ method: 'GET', mode: 'no-cors', url });

test('the worker serves a saved file only while it is the one the build shipped', async () => {
	const { run, stores, fetches } = worker();
	await run('install');
	await run('activate');
	const [name] = [...stores.keys()];
	const store = stores.get(name);
	const suite = [...store.keys()].find((href) => /suite\.js\?v=/u.test(href));
	assert.ok(suite, 'suite.js is saved at install');

	fetches.length = 0;
	const genuine = await run('fetch', { request: request(suite) });
	assert.equal(await genuine.text(), read('suite.js').toString());
	assert.deepEqual(fetches, [], 'a genuine copy comes from the cache');

	store.set(suite, new Response('globalThis.changed = true;'));
	const served = await run('fetch', { request: request(suite) });
	assert.equal(await served.text(), read('suite.js').toString());
	assert.deepEqual(fetches, [suite], 'a changed copy is fetched again');
	assert.equal(await store.get(suite).text(), read('suite.js').toString(), 'and replaced');
});

test('a changed saved shell is not served when the network is down', async () => {
	const online = worker();
	await online.run('install');
	const store = [...online.stores.values()][0];
	const page = `${ORIGIN}index.html`;
	store.set(page, new Response('<meta name="ooxml-root"><script>changed()</script>'));
	const offline = worker({ offline: true, stores: online.stores });
	const served = await offline
		.run('fetch', { request: { method: 'GET', mode: 'navigate', url: page } })
		.catch((error) => error);
	assert.ok(served instanceof TypeError, 'the browser shows its own offline page');
	assert.equal(store.has(page), false, 'the changed shell is deleted');

	await online.run('install');
	const restored = await worker({ offline: true, stores: online.stores }).run('fetch', {
		request: { method: 'GET', mode: 'navigate', url: page },
	});
	assert.equal(await restored.text(), read('index.html').toString(), 'a genuine shell still loads');
});

test('an older build cache is dropped, since its files cannot be checked', async () => {
	const { run, stores } = worker();
	stores.set(`ooxml-app-/ooxml/-older`, new Map([[`${ORIGIN}suite.js?v=old`, new Response('x')]]));
	await run('install');
	await run('activate');
	assert.deepEqual([...stores.keys()].filter((key) => key.startsWith('ooxml-app-')).length, 1);
});
