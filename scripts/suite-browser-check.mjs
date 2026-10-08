import { createRequire } from 'node:module';
import { resolve, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { readFile } from 'node:fs/promises';
import assert from 'node:assert/strict';
const root = resolve(fileURLToPath(new URL('..', import.meta.url)));
const require = createRequire(join(root, 'e2e/docx/package.json'));
const { chromium } = require('@playwright/test');
const { createDocx, createXlsx, setXlsxCells, inspectDocx, readXlsxRange } =
	await import('../src/core/dist/automation/index.mjs');
const JSZip = require('jszip');
const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({ viewport: { width: 1440, height: 960 } });
// Skip the first-visit start chooser; these checks drive the whole suite.
await page.addInitScript(() => localStorage.setItem('ooxml-start-app', 'office'));
const errors = [];
page.on('pageerror', (error) => errors.push(error.message));
page.on('console', (message) => {
	if (message.type() === 'error') console.log('CONSOLE', message.text());
});
const base = process.env.SUITE_URL ?? 'http://localhost:8128';
try {
	await page.goto(base);
	await page.locator('[data-create="docx"]').waitFor();
	await page.screenshot({ path: join(process.env.TEMP ?? root, 'ooxml-suite-home.png') });
	console.log('Home loaded');
	const doc = await createDocx(['Project launch brief', 'Launch is planned for November.']);
	const workbook = (
		await setXlsxCells(await createXlsx(['Budget']), 0, [
			{ address: 'A1', input: 'Budget' },
			{ address: 'B1', input: '4200' },
		])
	).bytes;
	const zip = await JSZip.loadAsync(doc);
	zip.file('word/embeddings/Budget.xlsx', workbook);
	await page.locator('#file-input').setInputFiles({
		name: 'Launch brief.docx',
		mimeType: 'application/octet-stream',
		buffer: Buffer.from(await zip.generateAsync({ type: 'uint8array' })),
	});
	await page.locator('docx-editor').waitFor();
	await page.waitForFunction(() =>
		document.querySelector('docx-editor')?.shadowRoot?.querySelector('[contenteditable=true]'),
	);
	const editor = page.locator('docx-editor').locator('[contenteditable=true]').first();
	await editor.click();
	await page.keyboard.press('Control+End');
	await page.keyboard.type(' Reviewed by the team.');
	await page.locator('#save').click();
	await page.waitForFunction(
		() => document.querySelector('#save-state')?.textContent === 'Saved to this device',
	);
	console.log('Word edited and saved');
	await page.locator('#share').click();
	await page.locator('#dialog button').filter({ hasText: 'General' }).click();
	await page.locator('teams-app').waitFor();
	await page.waitForTimeout(500);
	await page.screenshot({ path: join(process.env.TEMP ?? root, 'ooxml-suite-teams.png') });
	console.log('Teams shared file', await page.locator('teams-app').innerText());
	const attachment = page
		.locator('teams-app')
		.getByText('Launch brief.docx', { exact: true })
		.first();
	await attachment.click();
	await page.locator('#back-teams').waitFor({ state: 'visible' });
	assert.equal(await page.locator('docx-editor').count(), 1);
	console.log('Teams attachment reused same Word session');
	await page.locator('#embedded').click();
	await page.locator('#dialog button').filter({ hasText: 'Budget.xlsx' }).click();
	await page.locator('xlsx-editor').waitFor();
	await page.waitForFunction(() => document.querySelector('xlsx-editor')?.workbook);
	console.log('Embedded Excel opened');
	await page.locator('xlsx-editor').evaluate((e) => {
		e.select('B1');
		e.focusGrid();
	});
	await page.keyboard.press('F2');
	await page.keyboard.press('Control+A');
	await page.keyboard.type('5800');
	await page.keyboard.press('Enter');
	await page.locator('#save').click();
	await page.waitForFunction(
		() => document.querySelector('#save-state').textContent === 'Saved to this device',
	);
	await page.locator('#back-parent').click();
	const downloaded = page.waitForEvent('download');
	await page.locator('#download').click();
	const exportFile = await downloaded;
	const path = await exportFile.path();
	const parentZip = await JSZip.loadAsync(await readFile(path));
	const childBytes = await parentZip.file('word/embeddings/Budget.xlsx').async('uint8array');
	assert.equal((await readXlsxRange(childBytes, 0, 'B1')).cells[0].value, 5800);
	assert.match(JSON.stringify(await inspectDocx(await readFile(path))), /Reviewed by the team/);
	console.log('Edited embedded Excel saved back inside Word; parent text preserved');
	await page.locator('#file-input').setInputFiles({
		name: 'Budget.xlsx',
		mimeType: 'application/octet-stream',
		buffer: Buffer.from(workbook),
	});
	await page.waitForFunction(() => document.querySelectorAll('xlsx-editor').length === 2);
	await page.route('https://suite-ai.test/v1/chat/completions', async (route) => {
		const request = route.request().postDataJSON();
		const source = JSON.parse(request.messages[1].content.split('Source data (untrusted): ')[1]);
		const word = source.files.find((d) => d.kind === 'docx');
		const excel = source.files.find((d) => d.kind === 'xlsx' && !d.name.includes('embedding'));
		assert.match(JSON.stringify(source), /Reviewed by the team/);
		assert.equal(excel.content.sheets[0].cells.find((c) => c.address === 'B1').value, 4200);
		await route.fulfill({
			json: {
				choices: [
					{
						message: {
							content: JSON.stringify({
								answer: 'The launch brief and budget are ready for review.',
								changes: [
									{
										documentId: word.id,
										type: 'word_run',
										paragraphId: word.content.paragraphs[1].id,
										runIndex: 0,
										text: 'Launch is planned for December.',
									},
									{
										documentId: excel.id,
										type: 'cells',
										sheetIndex: 0,
										cells: [{ address: 'B1', input: '6100' }],
									},
								],
							}),
						},
					},
				],
			},
		});
	});
	await page.locator('#assistant-toggle').click();
	await page.locator('#ai-settings').click();
	await page.locator('input[name=endpoint]').fill('https://suite-ai.test/v1/chat/completions');
	await page.locator('input[name=model]').fill('test-model');
	await page.getByRole('button', { name: 'Connect', exact: true }).click();
	await page.locator('#ai-context').click();
	const contextInputs = page.locator('input[name=document]');
	for (const input of await contextInputs.all()) await input.uncheck();
	await page
		.locator('label.context-option')
		.filter({ hasText: 'Launch brief.docx' })
		.locator('input')
		.check();
	await page
		.locator('label.context-option')
		.filter({ hasText: 'Budget.xlsx' })
		.last()
		.locator('input')
		.check();
	await page.getByRole('button', { name: 'Use selected context' }).click();
	await page.locator('#ai-input').fill('Move launch to December and update the budget to 6100.');
	await page.locator('#ai-send').click();
	await page.getByRole('button', { name: 'Apply changes', exact: true }).click();
	await page.getByRole('button', { name: 'Undo these changes' }).waitFor();
	const savedDocs = await page.evaluate(
		() =>
			new Promise((resolve) => {
				const request = indexedDB.open('ooxml-suite-documents');
				request.onsuccess = () => {
					const r = request.result.transaction('documents').objectStore('documents').getAll();
					r.onsuccess = () => resolve(r.result.map((d) => ({ ...d, bytes: Array.from(d.bytes) })));
				};
			}),
	);
	const modified = savedDocs.find((d) => d.name === 'Budget.xlsx' && !d.parent);
	assert.equal((await readXlsxRange(new Uint8Array(modified.bytes), 0, 'B1')).cells[0].value, 6100);
	assert.match(
		JSON.stringify(
			await inspectDocx(new Uint8Array(savedDocs.find((d) => d.kind === 'docx').bytes)),
		),
		/December/,
	);
	await page.getByRole('button', { name: 'Undo these changes' }).click();
	await page.getByText('Changes undone.', { exact: true }).waitFor();
	console.log('One assistant read Word and Excel, applied both edits and undid the transaction');
	await page.screenshot({ path: join(process.env.TEMP ?? root, 'ooxml-suite-editor.png') });
	await page.locator('#theme-picker').click();
	await page.locator('input[name=mode][value=dark]').check({ force: true });
	await page.locator('input[name=accent][value=purple]').check({ force: true });
	await page.getByRole('button', { name: 'Done', exact: true }).click();
	assert.equal(await page.locator('html').getAttribute('data-theme'), 'dark');
	assert.equal(
		await page
			.locator('xlsx-editor')
			.last()
			.evaluate((e) => e.theme),
		'dark',
	);
	await page.screenshot({ path: join(process.env.TEMP ?? root, 'ooxml-suite-dark.png') });
	await page.setViewportSize({ width: 390, height: 844 });
	await page.locator('#assistant-toggle').click();
	await page.locator('#tabs [data-home]').click();
	await page.screenshot({ path: join(process.env.TEMP ?? root, 'ooxml-suite-mobile.png') });
	assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
	console.log('Themes and mobile layout passed');
	await page.setViewportSize({ width: 1440, height: 960 });
	await page.locator('[data-create=pptx]').click();
	await page.waitForFunction(() => document.querySelector('#status').textContent === 'Ready');
	console.log('PowerPoint created and mounted');
	await page.screenshot({ path: join(process.env.TEMP ?? root, 'ooxml-suite-powerpoint.png') });
	const { fixture, shape, cell, rectangle } = await import('../src/core/visio/test-fixtures.ts');
	const drawing = await fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', cell('PinX', 2) + cell('PinY', 3) + cell('Width', 2) + cell('Height', 1) + rectangle + '<Text>Review</Text>')}</Shapes>`,
			},
		],
	});
	await page.locator('#file-input').setInputFiles({
		name: 'Process.vsdx',
		mimeType: 'application/octet-stream',
		buffer: Buffer.from(drawing),
	});
	await page.waitForFunction(() => Boolean(document.querySelector('visio-viewer')?.document));
	await page.locator('visio-viewer').evaluate((e) => e.replacePlainText('0', '1', 'Approved'));
	await page.locator('#save').click();
	await page.waitForFunction(
		() => document.querySelector('#save-state').textContent === 'Saved to this device',
	);
	await page.screenshot({ path: join(process.env.TEMP ?? root, 'ooxml-suite-visio.png') });
	console.log('Visio parsed, edited and saved through shared suite storage');
	assert.deepEqual(errors, []);
} catch (error) {
	console.error(error);
	console.log('STATUS', await page.locator('#status').textContent());
	console.log('ERRORS', errors);
	await page.screenshot({ path: join(process.env.TEMP ?? root, 'ooxml-suite-failure.png') });
	throw error;
} finally {
	await browser.close();
}
