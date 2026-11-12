/**
 * Simple view, the suite's default layout: a document opens under one short command bar with
 * its app's ribbon held back, the file commands sit in the More menu, PowerPoint's slide
 * navigation and undo work from the bar, the ribbon comes back without reloading the deck, and
 * PowerPoint's own Ribbon Display Options are remembered by the suite.
 *
 * Run against a served build: bun run build:suite, bun run preview:suite, then
 * SUITE_URL=http://localhost:8123 node scripts/suite-simple-view-check.mjs
 */
import { createRequire } from 'node:module';
import assert from 'node:assert/strict';
const require = createRequire(new URL('../e2e/docx/package.json', import.meta.url));
const { chromium } = require('@playwright/test');
const base = process.env.SUITE_URL || 'http://127.0.0.1:8123';
const deck = new URL('../e2e/pptx/fixtures/slide-jump-links.pptx', import.meta.url);

const browser = await chromium.launch();
const context = await browser.newContext({ viewport: { width: 1440, height: 960 } });
// Only the start screen is skipped: the layout stays the default a new visitor gets.
await context.addInitScript(() => (globalThis.ooxmlSkipStart = true));
const page = await context.newPage();
page.setDefaultTimeout(30000);
const errors = [];
page.on('pageerror', (error) => errors.push(error.message));

const ready = () =>
	page.waitForFunction(() =>
		/^(Ready|Local|Connected)/u.test(document.getElementById('status')?.textContent ?? ''),
	);
const openDeck = async () => {
	await page.setInputFiles(
		'#file-input',
		decodeURIComponent(deck.pathname).replace(/^\/(\w:)/u, '$1'),
	);
	await page.waitForSelector('.document-host:not([hidden]) .pptxv');
	await ready();
};
const ribbon = () => page.locator('.document-host:not([hidden]) .pptxv-ribbon');

try {
	await page.goto(`${base}/?suite`);
	await page.waitForSelector('#file-input', { state: 'attached' });
	await openDeck();

	// The default: the bar, no ribbon, the file name with its extension apart.
	assert.equal(await page.evaluate(() => document.body.dataset.layout), 'simple');
	assert.equal(await ribbon().count(), 0, 'Simple view shows a ribbon');
	assert.match(await page.locator('#doc-title').innerText(), /slide-jump-links\s*\.pptx/u);
	assert.equal(await page.locator('#share').isVisible(), false, 'Share is not in the More menu');
	console.log('A deck opens in Simple view by default');

	// The file commands live in the More menu.
	await page.click('#actions-more');
	const items = await page.locator('#actions-menu [role="menuitem"]').allInnerTexts();
	for (const label of ['Rename', 'Download', 'Versions'])
		assert.ok(items.includes(label), `More menu lacks ${label}`);
	await page.keyboard.press('Escape');

	// Slide navigation and undo from the bar.
	const position = () => page.locator('#doc-position').innerText();
	const total = Number(/\/ (\d+)$/u.exec(await position())[1]);
	assert.ok(total > 1, 'the fixture needs more than one slide');
	await page.click('#doc-next');
	await page.waitForFunction(() =>
		/^Slide 2 \//u.test(document.getElementById('doc-position').textContent),
	);
	assert.equal(await page.locator('#doc-undo').isDisabled(), true);
	await page
		.locator('.document-host:not([hidden])')
		.getByText('Add Slide', { exact: true })
		.click();
	await page.waitForFunction(
		(count) => document.getElementById('doc-position').textContent.endsWith(`/ ${count}`),
		total + 1,
	);
	assert.equal(await page.locator('#doc-undo').isDisabled(), false);
	console.log('Slide navigation and undo state follow the deck');

	// Show ribbon: the same deck, its unsaved slide and its undo history stay.
	await page.click('#layout-toggle');
	await ribbon().waitFor();
	assert.equal(await page.evaluate(() => document.body.dataset.layout), 'office');
	assert.match(await page.locator('#save-state').innerText(), /Unsaved/u);
	await page.click('#layout-toggle');
	await ribbon().waitFor({ state: 'detached' });
	assert.match(await position(), new RegExp(`/ ${total + 1}$`, 'u'));
	await page.click('#doc-undo');
	await page.waitForFunction(
		(count) => document.getElementById('doc-position').textContent.endsWith(`/ ${count}`),
		total,
	);
	console.log('The ribbon comes and goes without reloading the deck');

	// PowerPoint's Ribbon Display Options: the suite remembers "Show tabs only"; Full-screen
	// mode is Simple view.
	await page.click('#layout-toggle');
	await ribbon().waitFor();
	const choose = async (mode) => {
		await page.click('.document-host:not([hidden]) .pptxv-ribbon-display-btn');
		await page.click(
			`.document-host:not([hidden]) [role="menuitemradio"][data-ribbon-display="${mode}"]`,
		);
	};
	await choose('tabs');
	await page.waitForFunction(
		() => JSON.parse(localStorage.getItem('ooxml-suite-settings')).ribbonDisplay === 'tabs',
	);
	assert.equal(await ribbon().getAttribute('data-ribbon-display'), 'tabs');
	await choose('fullScreen');
	await page.waitForFunction(() => document.body.dataset.layout === 'simple');
	assert.equal(await ribbon().count(), 0);
	await page.click('#layout-toggle');
	await ribbon().waitFor();
	assert.equal(await ribbon().getAttribute('data-ribbon-display'), 'tabs');
	console.log('PowerPoint ribbon options are remembered; full-screen mode is Simple view');

	assert.deepEqual(errors, []);
} finally {
	await browser.close();
}
