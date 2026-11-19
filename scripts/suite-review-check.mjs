// Browser check: a document showing the assistant's suggestions never leaves the suite before
// they are applied. For Word, Excel, PowerPoint and Visio it opens a review preview, then checks
// that the suite's Download and Ctrl+P refuse with a message, that the editor's File tab offers
// none of Save As, Export, Print and Share (Visio's Save As saves through the suite, which
// refuses), that its Copy and export commands are gone, that no download starts, and that the
// commands return once the preview is discarded.
// Run against a preview of suite-dist:
//   SUITE_URL=http://localhost:8123 node scripts/suite-review-check.mjs
import { createRequire } from 'node:module';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';
import { ribbonLayoutInit } from './suite-check-init.mjs';
const root = resolve(fileURLToPath(new URL('..', import.meta.url)));
const require = createRequire(join(root, 'e2e/docx/package.json'));
const { chromium } = require('@playwright/test');
const { createDocx, createXlsxFromSheets, editVisio } =
	await import('../src/core/dist/automation/index.mjs');
const { createVsdx } = await import('../src/core/dist/visio/index.mjs');
const base = process.env.SUITE_URL ?? 'http://localhost:8123';
const LEAVING = ['saveAs', 'save-as', 'export', 'print', 'share'];
const browser = await chromium.launch({ headless: true });
const errors = [];
try {
	const context = await browser.newContext({
		viewport: { width: 1280, height: 900 },
		acceptDownloads: true,
	});
	await context.addInitScript(ribbonLayoutInit);
	const page = await context.newPage();
	page.on('pageerror', (error) => errors.push(error.message));
	const downloads = [];
	page.on('download', (download) => downloads.push(download.suggestedFilename()));
	await page.goto(base);
	await page.locator('#file-input').waitFor({ state: 'attached' });
	const pane = page.locator('#assistant');
	const panel = page.locator('.review-panel');

	/** Asks the assistant to title-case document `name` and shows its suggestions in place. */
	const review = async (name) => {
		if (await pane.isHidden()) await page.locator('#assistant-toggle').click();
		await page.locator('#ai-new').click();
		// Closing and opening the pane attaches the document on screen.
		await page.locator('#assistant-toggle').click();
		await page.locator('#assistant-toggle').click();
		await page.locator('#ai-scope .chat-chip', { hasText: name }).waitFor();
		// Only this document: the ones checked before stay attached to the new chat.
		await page.waitForFunction((name) => {
			const other = [...document.querySelectorAll('#ai-scope .chat-chip')].find(
				(chip) => !chip.textContent.includes(name),
			);
			other?.querySelector('[data-detach]')?.click();
			return !other;
		}, name);
		await page.locator('#ai-input').fill('Title-case this document');
		await page.locator('#ai-send').click();
		await pane.getByRole('button', { name: 'Review in document' }).click();
		await panel.waitFor();
		await panel.getByText('Save As, Export, Download, Print, Share and Copy are off').waitFor();
	};
	const open = async (name, bytes, selector) => {
		await page.locator('#rail [data-home]').click();
		await page.locator('#file-input').setInputFiles({
			name,
			mimeType: 'application/octet-stream',
			buffer: Buffer.from(bytes),
		});
		await page.locator(selector).last().waitFor();
	};
	/** The suite's own commands refuse with a message naming the document. */
	const suiteRefuses = async (selector) => {
		for (const run of [
			() => page.evaluate(() => document.getElementById('download').click()),
			async () => {
				// From inside the editor, where its own shortcuts would take the key.
				await page
					.locator(selector)
					.last()
					.click({ position: { x: 60, y: 300 } });
				await page.keyboard.press('Control+p');
			},
		]) {
			await page.evaluate(() => {
				document.getElementById('status').textContent = '';
			});
			await run();
			await page.locator('#status', { hasText: "showing the assistant's suggestions" }).waitFor();
		}
	};
	/** The File tab entries of Word, Excel and Visio (the shared backstage's ids). */
	const backstageIds = (selector) =>
		page.evaluate(
			(selector) =>
				[...document.querySelectorAll(selector)]
					.at(-1)
					.shadowRoot.querySelector('office-ui-backstage')
					.items.map((item) => item.id),
			selector,
		);
	/** Commands the ribbon, menus, dialogs and File tab pages still offer, by id. */
	const shown = (selector, ids) =>
		page.evaluate(
			({ selector, ids }) => {
				const host = [...document.querySelectorAll(selector)].at(-1);
				return ids.filter((id) =>
					[...host.shadowRoot.querySelectorAll(`[command="${id}"]`)].some(
						(element) => !element.hidden,
					),
				);
			},
			{ selector, ids },
		);

	for (const product of [
		{
			name: 'Word',
			selector: 'docx-editor',
			file: ['Plain brief.docx', await createDocx(['launch brief', 'ships in november.'])],
			end: 'Reject all',
			leaving: ['saveAs', 'export', 'print'],
		},
		{
			name: 'Excel',
			selector: 'xlsx-editor',
			file: [
				'Budget.xlsx',
				await createXlsxFromSheets([{ name: 'Budget', rows: [['launch budget', 42]] }]),
			],
			end: 'Discard',
			leaving: ['saveAs', 'export', 'print', 'share'],
		},
		{
			name: 'Visio',
			selector: 'visio-viewer',
			file: [
				'Flow.vsdx',
				(
					await editVisio(await createVsdx(), [
						{
							type: 'create-text-box',
							pageId: '0',
							shapeId: '1',
							x: 2,
							y: 2,
							width: 3,
							height: 1,
							text: 'process start',
						},
					])
				).bytes,
			],
			end: 'Discard',
			leaving: ['export', 'print', 'share'],
			// The Export page holds the PDF, PNG and SVG exports; these live on the ribbon and dialogs.
			commands: ['shape-report-export-csv', 'shape-report-copy', 'copy', 'cut'],
		},
	]) {
		await open(...product.file, product.selector);
		const before = downloads.length;
		await review(product.file[0]);
		await suiteRefuses(product.selector);
		const ids = await backstageIds(product.selector);
		assert.deepEqual(
			ids.filter((id) => LEAVING.includes(id) && id !== 'save-as'),
			[],
			`${product.name} offers no command that takes the document out`,
		);
		if (product.commands) assert.deepEqual(await shown(product.selector, product.commands), []);
		if (product.name === 'Visio') {
			// Visio's Save As page downloads a copy through the suite's Save, which refuses.
			await page.evaluate(() => {
				document.getElementById('status').textContent = '';
				[...document.querySelectorAll('visio-viewer')]
					.at(-1)
					.shadowRoot.querySelector('[data-backstage-action="download"]')
					.click();
			});
			await page.locator('#status', { hasText: "showing the assistant's suggestions" }).waitFor();
		}
		assert.equal(downloads.length, before, `${product.name} downloaded nothing`);
		await panel.getByRole('button', { name: product.end }).click();
		await panel.waitFor({ state: 'detached' });
		const after = await backstageIds(product.selector);
		for (const id of product.leaving)
			assert.ok(after.includes(id), `${product.name} offers ${id} again after the review`);
		if (product.commands)
			assert.deepEqual(await shown(product.selector, product.commands), product.commands);
		console.log(`${product.name}: the preview offered no way out; the commands came back after`);
	}

	// PowerPoint: a deck the assistant creates, then title-cases.
	if (await pane.isHidden()) await page.locator('#assistant-toggle').click();
	await page.locator('#ai-new').click();
	await page.locator('#ai-input').fill('Create a three-slide presentation about the launch plan');
	await page.locator('#ai-send').click();
	await pane.getByRole('button', { name: /^Create/u }).click();
	await page.locator('.pptxv').last().waitFor();
	const before = downloads.length;
	await review('.pptx');
	await suiteRefuses('.pptxv');
	const navLabels = async () => {
		const viewer = page.locator('.pptxv').last();
		await viewer.getByRole('tab', { name: 'File', exact: true }).first().click();
		const nav = viewer.locator('[data-pptx-backstage-nav-item]');
		await nav.first().waitFor();
		const labels = (await nav.allTextContents()).map((label) => label.trim());
		await page.keyboard.press('Escape');
		return labels;
	};
	const leaving = ['Save As', 'Export', 'Print', 'Share'];
	assert.deepEqual(
		(await navLabels()).filter((label) => leaving.includes(label)),
		[],
		'PowerPoint offers no command that takes the deck out',
	);
	assert.equal(downloads.length, before, 'PowerPoint downloaded nothing');
	await panel.getByRole('button', { name: 'Discard' }).click();
	await panel.waitFor({ state: 'detached' });
	const back = await navLabels();
	for (const label of ['Save As', 'Export', 'Print'])
		assert.ok(back.includes(label), `PowerPoint offers ${label} again after the review`);
	console.log('PowerPoint: the preview offered no way out; the commands came back after');
	assert.deepEqual(errors, []);
	await context.close();
} finally {
	await browser.close();
}
