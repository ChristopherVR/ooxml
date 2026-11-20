/**
 * What the suite's screens download: the suite is built into a directory, served, and each screen
 * is opened in Chromium; a screen's size is the gzipped size of every JavaScript file it requested,
 * each counted once. Reading the requests, not the chunk graph, keeps a lazy import a screen makes
 * while it opens (an editor's loader, the automation bundle behind New) in its number.
 */
import { spawn } from 'node:child_process';
import { readdirSync, readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { gzipSync } from 'node:zlib';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const LANDING = '#create-new svg';
const create = (kind, ready) => ({
	path: '/?suite#/',
	ready: LANDING,
	create: kind,
	opened: ready,
});

/**
 * Each measured screen: the page, what shows it is up, and for New the file type and the editor
 * that must open. Every page but the start chooser skips the chooser, as a returning user does.
 */
export const SUITE_ROUTES = {
	'start chooser': { path: '/', ready: '.start-center' },
	home: { path: '/?suite#/', ready: LANDING },
	'word landing': { path: '/apps/word/', ready: LANDING },
	'excel landing': { path: '/apps/excel/', ready: LANDING },
	'powerpoint landing': { path: '/apps/powerpoint/', ready: LANDING },
	'visio landing': { path: '/apps/visio/', ready: LANDING },
	chat: { path: '/?suite', ready: '#chat textarea, #chat-input' },
	'new word': create('docx', 'docx-editor'),
	'new excel': create('xlsx', 'xlsx-editor'),
	'new powerpoint': create('pptx', '.pptxv, [class*=pptx]'),
	'new visio': create('vsdx', 'visio-viewer'),
	teams: { path: '/apps/teams/', ready: 'teams-app' },
};
/** A screen has settled once it has requested no script for this long. */
const QUIET_MS = 1500;
const TIMEOUT_MS = 90_000;

export const gzipBytes = (bytes) => gzipSync(bytes, { level: 9 }).length;

/** Every script file under `dir`, leaving out the paths `skip` accepts. */
export function* scriptFiles(dir, skip = () => false) {
	for (const entry of readdirSync(dir, { withFileTypes: true })) {
		const path = join(dir, entry.name);
		if (skip(path)) continue;
		if (entry.isDirectory()) yield* scriptFiles(path, skip);
		else if (/\.[cm]?js$/.test(entry.name)) yield path;
	}
}

/** `node scripts/preview-site.mjs` on a free port, serving `site`. */
async function serve(site) {
	const server = spawn(process.execPath, [join(ROOT, 'scripts/preview-site.mjs'), '0', site], {
		stdio: ['ignore', 'pipe', 'inherit'],
	});
	const url = await new Promise((done, fail) => {
		server.once('exit', (code) => fail(new Error(`The preview server exited (${code}).`)));
		server.stdout.on('data', (data) => {
			const port = /localhost:(\d+)/.exec(String(data))?.[1];
			if (port) done(`http://127.0.0.1:${port}`);
		});
	});
	return { url, close: () => server.kill() };
}

/** The script files `route` requests, from opening the page until it has settled. */
async function requested(browser, url, route) {
	const context = await browser.newContext({ serviceWorkers: 'block' });
	try {
		if (route.path !== '/')
			await context.addInitScript(() => {
				globalThis.ooxmlSkipStart = true;
			});
		const page = await context.newPage();
		page.setDefaultTimeout(TIMEOUT_MS);
		const files = new Set();
		const errors = [];
		let last = Date.now();
		page.on('request', (request) => {
			const path = new URL(request.url()).pathname;
			if (/\.[cm]?js$/.test(path)) {
				files.add(decodeURIComponent(path.slice(1)));
				last = Date.now();
			}
		});
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(url + route.path, { waitUntil: 'load' });
		await page.locator(route.ready).first().waitFor();
		if (route.create) {
			await page.locator(`#create-row button[data-create="${route.create}"]`).click();
			await page.locator(route.opened).first().waitFor({ state: 'attached' });
			await page.waitForFunction(() => document.querySelector('#status')?.textContent === 'Ready');
		}
		// Late lazy loads (an idle prefetch, a ribbon built after the first paint) still count.
		const deadline = Date.now() + TIMEOUT_MS;
		while (Date.now() - last < QUIET_MS && Date.now() < deadline) await page.waitForTimeout(250);
		if (errors.length) throw new Error(`The page threw: ${errors.join('; ')}`);
		return [...files].sort();
	} finally {
		await context.close();
	}
}

/**
 * `{ "suite/...": bytes }` for a built suite in `site`: all of its JavaScript (the docs pages
 * aside), and per route what Chromium requested. `onRoute` sees each route's files.
 */
export async function measureSuiteSite(site, { onRoute = () => {} } = {}) {
	const docs = join(site, 'docs');
	const gzip = new Map();
	for (const file of scriptFiles(site, (path) => path === docs))
		gzip.set(relative(site, file).replaceAll('\\', '/'), gzipBytes(readFileSync(file)));
	const sizes = { 'suite/all javascript': [...gzip.values()].reduce((a, b) => a + b, 0) };
	// The workspace declares Playwright for its browser tests; the accessibility gates use Chromium.
	const { chromium } = createRequire(join(ROOT, 'e2e/a11y/package.json'))('@playwright/test');
	const server = await serve(site);
	const browser = await chromium.launch();
	try {
		for (const [name, route] of Object.entries(SUITE_ROUTES)) {
			let files;
			try {
				files = await requested(browser, server.url, route);
			} catch (error) {
				throw new Error(`Route ${name} (${route.path}) did not open: ${error.message}`, {
					cause: error,
				});
			}
			const unknown = files.filter((file) => !gzip.has(file));
			if (unknown.length) throw new Error(`Route ${name} requested missing files: ${unknown}`);
			sizes[`suite/route/${name}`] = files.reduce((sum, file) => sum + gzip.get(file), 0);
			onRoute(name, files);
		}
	} finally {
		await browser.close();
		server.close();
	}
	return sizes;
}
