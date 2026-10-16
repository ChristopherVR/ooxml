/**
 * Visio Presentation Mode inside the suite: switching to another suite tab while presenting
 * must leave no fullscreen element and nothing painted over the new tab. The Fullscreen API
 * keeps a promoted element above the whole page, so an editor hidden behind a tab switch has to
 * give it up (apps/office-suite/src/main.js, src/ui/src/visio/viewer-presentation.ts).
 *
 * Run against a served build: bun run build:suite, bun run preview:suite, then
 * SUITE_URL=http://localhost:8171 node scripts/suite-visio-presentation-check.mjs
 */
import { createRequire } from 'node:module';
import assert from 'node:assert/strict';
const require = createRequire(new URL('../e2e/docx/package.json', import.meta.url));
const { chromium, expect } = require('@playwright/test');
const base = process.env.SUITE_URL || 'http://127.0.0.1:8171';
const browser = await chromium.launch();
const context = await browser.newContext({ viewport: { width: 1440, height: 960 } });
// Skip the start chooser; this check drives the whole suite.
await context.addInitScript(() => (globalThis.ooxmlSkipStart = true));
const page = await context.newPage();
page.setDefaultTimeout(20000);
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));

/** What a click at the middle of the page would hit, and whether anything is fullscreen. */
const paint = () =>
	page.evaluate(() => {
		const hit = document.elementFromPoint(innerWidth / 2, innerHeight / 2);
		return {
			fullscreen: document.fullscreenElement?.localName ?? null,
			insideVisio: Boolean(hit?.closest('visio-viewer')),
			hit: hit?.id || hit?.localName || null,
		};
	});

try {
	await page.goto(base);
	await page.locator('#rail [data-home]').click();
	await page.locator('[data-create=vsdx]').click();
	const visio = page.locator('visio-viewer');
	await expect(visio.locator('office-ui-ribbon.toolbar')).toBeVisible();
	await page.waitForFunction(() => Boolean(document.querySelector('visio-viewer')?.document));

	// F5 is a real key press, so the browser grants the user activation fullscreen needs.
	await visio.locator('.viewport').focus();
	await page.keyboard.press('F5');
	const presenting = () =>
		visio.evaluate((el) => {
			const overlay = el.shadowRoot.querySelector('.presentation');
			return Boolean(overlay) && !overlay.hidden;
		});
	await expect.poll(presenting).toBe(true);
	// The request resolves a moment after the key press; the check means nothing without it.
	await expect.poll(async () => (await paint()).fullscreen).toBe('visio-viewer');
	const entered = await paint();
	assert.equal(entered.insideVisio, true, 'the presentation covers the page while it runs');
	console.log('Presentation started in fullscreen.');

	// Switch to the Home tab, then to a second document, both while presenting.
	await page.evaluate(() => {
		// Tab buttons sit under the overlay; a script click is what a keyboard tab switch does.
		document.querySelector('#tabs [data-home]')?.click();
	});
	await expect.poll(async () => (await paint()).fullscreen).toBe(null);
	const home = await paint();
	assert.equal(home.insideVisio, false, 'nothing of the Visio editor paints over Home');
	await expect(page.locator('#home')).toBeVisible();
	assert.equal(await visio.evaluate((el) => el.offsetParent === null), true);

	// Back to the drawing: Presentation Mode ended with the tab switch, the editor is usable.
	await page.locator('#tabs [data-open]').first().click();
	await expect(visio.locator('office-ui-ribbon.toolbar')).toBeVisible();
	const back = await paint();
	assert.equal(back.fullscreen, null, 'no fullscreen element after returning');
	assert.equal(await presenting(), false, 'Presentation Mode is not left running behind the tab');

	// A second document tab takes over while presenting.
	await visio.locator('.viewport').focus();
	await page.keyboard.press('F5');
	await expect.poll(presenting).toBe(true);
	await expect.poll(async () => (await paint()).fullscreen).toBe('visio-viewer');
	await page.evaluate(() => document.querySelector('#tabs [data-home]')?.click());
	await page.locator('[data-create=pptx]').click();
	await page.waitForFunction(() => document.querySelector('#status').textContent === 'Ready');
	const other = await paint();
	assert.equal(other.fullscreen, null, 'no fullscreen element over the new PowerPoint tab');
	assert.equal(other.insideVisio, false, 'nothing of the Visio editor paints over the new tab');
	assert.deepEqual(errors, []);
	console.log(
		'Visio Presentation Mode: tab switches leave no fullscreen element or overlay passed.',
	);
} finally {
	await browser.close();
}
