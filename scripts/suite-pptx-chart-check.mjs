/**
 * PowerPoint chart styling inside the suite: the Chart Design galleries apply, the
 * properties panel is laid out without help from a host CSS reset (the suite ships
 * none), and the edits survive Save, Download and reopening the file.
 *
 * Run against a served build: bun run build:suite, bun run preview:suite, then
 * SUITE_URL=http://localhost:8123 node scripts/suite-pptx-chart-check.mjs
 */
import { createRequire } from 'node:module';
import { join } from 'node:path';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
const require = createRequire(new URL('../e2e/docx/package.json', import.meta.url));
const { chromium, expect } = require('@playwright/test');
const base = process.env.SUITE_URL || 'http://127.0.0.1:8123';
const deck = new URL('../e2e/pptx/fixtures/chart-gallery.pptx', import.meta.url);
const browser = await chromium.launch();
const context = await browser.newContext({ viewport: { width: 1440, height: 960 } });
await context.addInitScript(() => (globalThis.ooxmlSkipStart = true));
const page = await context.newPage();
page.setDefaultTimeout(20000);
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));

// Reopening adds a second editor tab and hides the first, so only the visible one counts.
const stageChart = () =>
	page
		.locator('[data-pptx-viewport] [data-pptx-element="true"][aria-roledescription="chart"]')
		.filter({ visible: true })
		.first();
/** The fill of the chart's first plotted mark (series 0) on the main slide. */
const firstMarkFill = () =>
	page.evaluate(
		() =>
			[
				...document.querySelectorAll(
					'[aria-roledescription="slide"] [aria-roledescription="chart"] svg',
				),
			]
				.find((svg) => svg.checkVisibility())
				?.querySelector('[data-chart-part="dataPoint"][data-chart-series="0"]')
				?.getAttribute('fill') ?? null,
	);
async function openDeck(name, buffer) {
	await page
		.locator('#file-input')
		.setInputFiles({ name, mimeType: 'application/octet-stream', buffer });
	await stageChart().waitFor();
}
async function selectChart() {
	const box = await stageChart().boundingBox();
	await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
	await page.locator('[data-ribbon-contextual-tab="chartDesign"]').first().click();
}
const gallery = (name) =>
	page.locator(`pptx-ui-ribbon-gallery:has([data-ribbon-gallery="${name}"])`).first();

try {
	await page.goto(base);
	await openDeck('Chart styling.pptx', await readFile(deck));
	await selectChart();
	const authored = await firstMarkFill();

	// Change Colors: a monochromatic palette recolours the first series.
	await gallery('chartColors').locator('[data-ribbon-gallery]').click();
	await page.locator('[data-ribbon-gallery-popup] [data-gallery-item]').nth(5).click();
	await expect.poll(firstMarkFill).not.toBe(authored);
	const recoloured = await firstMarkFill();

	// The properties panel: every chart field on its own row, themed controls.
	const layout = await page
		.locator('.pptxv-inspector')
		.filter({ visible: true })
		.first()
		.evaluate((panel) => {
			const labels = [...panel.querySelectorAll('label')];
			const row = (caption) =>
				labels
					.find((label) => label.textContent?.trim().startsWith(caption))
					?.getBoundingClientRect();
			const title = row('Title');
			const type = row('Chart type');
			const input = labels.find((l) => l.textContent?.startsWith('Title'))?.querySelector('input');
			const textarea = panel.querySelector('textarea');
			const fieldset = panel.querySelector('fieldset');
			return {
				fieldsOnOwnRows: !!title && !!type && type.top >= title.bottom - 1,
				inputBorder: input ? getComputedStyle(input).borderTopStyle : null,
				textareaFont: textarea ? getComputedStyle(textarea).fontFamily : null,
				fieldsetBorder: fieldset ? getComputedStyle(fieldset).borderTopStyle : null,
				placeholders: /\{\{\w+\}\}/u.test(panel.textContent ?? ''),
			};
		});
	assert.equal(layout.fieldsOnOwnRows, true, 'chart fields run together on one line');
	assert.equal(layout.inputBorder, 'solid');
	assert.doesNotMatch(layout.textareaFont ?? '', /monospace/u);
	assert.notEqual(layout.fieldsetBorder, 'groove');
	assert.equal(layout.placeholders, false, 'a caption shows a raw {{placeholder}}');
	await page.screenshot({ path: join(process.env.TEMP ?? '.', 'suite-pptx-chart-styling.png') });

	// Save, download, reopen: the recoloured chart comes back.
	await page.locator('#save').click();
	await expect(page.locator('#save-state')).toHaveText('Saved to this device');
	const download = page.waitForEvent('download');
	await page.locator('#download').click();
	const saved = await readFile(await (await download).path());
	await openDeck('Chart styling reopened.pptx', saved);
	await expect.poll(firstMarkFill).toBe(recoloured);
	assert.deepEqual(errors, []);
	console.log(
		'PowerPoint charts: colours apply, panel lays out without a host reset, edits persist.',
	);
} finally {
	await browser.close();
}
