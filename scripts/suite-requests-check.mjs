/**
 * What each first screen and each New command downloads, read from the browser's request list
 * against a served build (`bun run preview:suite`, or SUITE_URL):
 * - the page links one stylesheet and every rule arrives with it (no `@import` chain);
 * - the home and the product landing pages load neither the Teams core nor the package tooling;
 * - New in each app loads neither the automation bundle nor the SmartArt definitions, which
 *   wait for an assistant request or a SmartArt.
 */
import { createRequire } from 'node:module';
import assert from 'node:assert/strict';
const require = createRequire(new URL('../e2e/docx/package.json', import.meta.url));
const { chromium } = require('@playwright/test');
const base = (process.env.SUITE_URL ?? 'http://127.0.0.1:8123').replace(/\/$/, '');

/** Chunks (by the module they are named after) a screen must not request. */
const TEAMS_CORE = /\/suite-assets\/teams-/;
const PACKAGE_TOOLING = /\/suite-assets\/(opc|jszip|yjs|xpath)-/;
const AUTOMATION = /\/suite-assets\/automation-/;
const SMARTART_DEFINITIONS = /\/suite-assets\/builtin-/;

const browser = await chromium.launch();
const failures = [];
/** Loads `path` in a fresh context, runs `then`, and returns every URL the page requested. */
async function requests(path, then = async () => {}) {
	const context = await browser.newContext({ serviceWorkers: 'block' });
	if (path !== '/')
		await context.addInitScript(() => {
			globalThis.ooxmlSkipStart = true;
		});
	const page = await context.newPage();
	page.setDefaultTimeout(60000);
	const urls = [];
	const errors = [];
	page.on('request', (request) => urls.push(request.url()));
	page.on('pageerror', (error) => errors.push(error.message));
	try {
		await page.goto(base + path, { waitUntil: 'load' });
		await then(page);
		// Late lazy loads (an idle prefetch, a ribbon built after the first paint) still count.
		await page.waitForTimeout(1500);
		assert.deepEqual(errors, [], `${path}: page errors`);
		return urls;
	} finally {
		await context.close();
	}
}
function check(name, urls, rules) {
	for (const [label, pattern] of rules) {
		const found = urls.filter((url) => pattern.test(url));
		if (found.length) failures.push(`${name} requested ${label}: ${found.join(', ')}`);
	}
	console.log(`${name}: ${urls.length} requests`);
}

try {
	const chooser = await requests('/', (page) => page.locator('.start-center').waitFor());
	const sheets = chooser.filter((url) => new URL(url).pathname.endsWith('.css'));
	if (sheets.length !== 1)
		failures.push(`the start page requested ${sheets.length} stylesheets: ${sheets.join(', ')}`);
	check('the start page', chooser, [['the Teams core', TEAMS_CORE]]);

	const landing = [
		['the home', '/?suite#/'],
		['the Word landing page', '/apps/word/'],
		['the Excel landing page', '/apps/excel/'],
		['the PowerPoint landing page', '/apps/powerpoint/'],
		['the Visio landing page', '/apps/visio/'],
	];
	for (const [name, path] of landing)
		check(name, await requests(path, (page) => page.locator('#create-new svg').waitFor()), [
			['the Teams core', TEAMS_CORE],
			['the package tooling', PACKAGE_TOOLING],
			['the automation bundle', AUTOMATION],
		]);

	const editors = {
		docx: 'docx-editor',
		xlsx: 'xlsx-editor',
		pptx: '.pptxv, [class*=pptx]',
		vsdx: 'visio-viewer',
	};
	for (const [kind, ready] of Object.entries(editors))
		check(
			`New ${kind}`,
			await requests('/?suite#/', async (page) => {
				await page.locator(`#create-row button[data-create="${kind}"]`).click();
				await page.locator(ready).first().waitFor({ state: 'attached' });
				await page.waitForFunction(
					() => document.querySelector('#status')?.textContent === 'Ready',
				);
			}),
			[
				['the automation bundle', AUTOMATION],
				['the SmartArt definitions', SMARTART_DEFINITIONS],
			],
		);
} finally {
	await browser.close();
}
assert.deepEqual(failures, [], failures.join('\n'));
console.log('Request checks passed');
