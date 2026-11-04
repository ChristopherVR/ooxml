/** Real suite/editor integration with HTTP doubles kept strictly in the test runner. */
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { readFile } from 'node:fs/promises';
import { interceptOneDrive } from './test-support/onedrive-http.mjs';
import { checkConcurrentMerges } from './test-support/onedrive-merge-checks.mjs';
import { checkCustomStorage } from './test-support/custom-storage-checks.mjs';

const require = createRequire(new URL('../e2e/docx/package.json', import.meta.url));
const { chromium, expect } = require('@playwright/test');
const { createDocx, createXlsx, inspectDocx, readXlsxRange, setXlsxCells } =
	await import('../src/core/dist/automation/index.mjs');
const { createPresentationFromOutline } = await import('../src/core/dist/pptx/index.mjs');
const { parseVsdx } = await import('../src/core/dist/visio/index.mjs');
const JSZip = require('jszip');
const { fixture, shape, cell, rectangle } = await import('../src/core/visio/test-fixtures.ts');
const fixtures = {
	docx: await createDocx(['Shared Word document.']),
	xlsx: await createXlsx(['Sheet1']),
	pptx: await createPresentationFromOutline([{ title: 'Shared PowerPoint' }]),
	vsdx: await fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', cell('PinX', 2) + cell('PinY', 3) + cell('Width', 2) + cell('Height', 1) + rectangle + '<Text>Shared Visio drawing.</Text>')}</Shapes>`,
			},
		],
	}),
};
const drawing = await parseVsdx(fixtures.vsdx);
const drawingPage = drawing.pages.find((p) => p.shapes.some((s) => s.text?.plainText?.trim()));
const drawingShape = drawingPage?.shapes.find((s) => s.text?.plainText?.trim());
assert.ok(drawingShape, 'the Visio fixture needs an editable text shape');
const base = process.env.SUITE_URL ?? 'http://127.0.0.1:8146';
const browser = await chromium.launch();
const context = await browser.newContext({
	viewport: { width: 1440, height: 1000 },
	serviceWorkers: 'block',
});
await context.addInitScript(() => {
	globalThis.ooxmlSkipStart = true;
	globalThis.ooxmlOneDriveOptions = {
		getAccessToken: async () => {
			if (globalThis.__oneDriveAuthFail) throw new Error('Private token provider failure');
			return 'test-only-token';
		},
		canWrite: () => !globalThis.__oneDriveReadOnly,
	};
});
const cloud = await interceptOneDrive(context, fixtures);
const page = await context.newPage();
page.setDefaultTimeout(30000);
const errors = [];
page.on('pageerror', (error) => errors.push(error.message));
page.on('requestfailed', (request) =>
	console.log('Request failed:', new URL(request.url()).hostname, request.failure()?.errorText),
);

async function openCloud(kind) {
	if (!(await page.locator('#open-onedrive').isVisible()))
		await page.locator('#rail [data-home]').click();
	await page.locator('#open-onedrive').click();
	await page.locator('#onedrive-link').fill(`https://tenant.sharepoint.com/links/${kind}`);
	await page.locator('#onedrive-open-submit').click();
	try {
		await expect(page.locator('#dialog')).toBeHidden();
	} catch (error) {
		console.log('Open status:', await page.locator(':light(#status)').textContent(), errors);
		throw error;
	}
	await expect(page.locator('#save-state')).toContainText('OneDrive');
}
async function edit(kind, value = 'Cloud edit') {
	const host = page.locator('.document-host:not([hidden])');
	if (kind === 'docx') {
		const text = host.locator('docx-editor').locator('[contenteditable=true]').first();
		await text.click();
		await page.keyboard.press('Control+End');
		await page.keyboard.type(` ${value}`);
	} else if (kind === 'xlsx') {
		await host.locator('xlsx-editor').evaluate((editor) => {
			editor.select('A1');
			editor.focusGrid();
		});
		await page.keyboard.press('F2');
		await page.keyboard.press('Control+A');
		await page.keyboard.type(value);
		await page.keyboard.press('Enter');
	} else if (kind === 'pptx') {
		await host.getByRole('button', { name: 'New Slide', exact: true }).click();
	} else {
		await host
			.locator('visio-viewer')
			.evaluate(
				(editor, target) => editor.replacePlainText(target.page, target.shape, target.value),
				{
					page: drawingPage.id,
					shape: drawingShape.id,
					value,
				},
			);
	}
}
async function nativeSave(kind) {
	const host = page.locator('.document-host:not([hidden])');
	if (kind === 'pptx')
		await host.getByRole('button', { name: 'Save', exact: true }).first().click();
	else
		await host
			.locator({ docx: 'docx-editor', xlsx: 'xlsx-editor', vsdx: 'visio-viewer' }[kind])
			.evaluate((editor) => {
				const event = new CustomEvent('file-command', {
					detail: { command: 'save' },
					cancelable: true,
					bubbles: true,
				});
				editor.dispatchEvent(event);
				if (!event.defaultPrevented)
					throw new Error('Suite did not handle the native Save contract');
			});
}
try {
	await page.goto(`${base}/#/`);
	for (const kind of ['docx', 'xlsx', 'pptx', 'vsdx']) {
		await openCloud(kind);
		await edit(kind);
		await nativeSave(kind);
		await expect.poll(() => cloud.writes.filter((id) => id === kind).length).toBe(1);
		await expect(page.locator('#save-state')).toContainText('Saved to OneDrive');
		assert.notDeepEqual(cloud.files.get(kind).bytes, Buffer.from(fixtures[kind]));
		await page.locator('#share').click();
		const share = page.locator('#onedrive-share-form');
		await share.getByRole('button', { name: 'Get OneDrive link', exact: true }).click();
		await expect(page.locator('#onedrive-share-link')).toHaveValue(/\/documents\//u);
		assert.equal(cloud.shares.length, ['docx', 'xlsx', 'pptx', 'vsdx'].indexOf(kind));
		await share.locator('select[name=type]').selectOption('view');
		await share.getByRole('button', { name: 'Get OneDrive link', exact: true }).click();
		await expect(page.locator('#onedrive-share-link')).toHaveValue(/test-capability/u);
		assert.deepEqual(cloud.shares.at(-1), { id: kind, type: 'view', scope: 'organization' });
		await page.locator('#dialog-close').click();
		console.log(`${kind}: opened and native Save updated the same OneDrive item`);
	}
	assert.match(JSON.stringify(await inspectDocx(cloud.files.get('docx').bytes)), /Cloud edit/u);
	assert.equal(
		(await readXlsxRange(cloud.files.get('xlsx').bytes, 0, 'A1')).cells[0].value,
		'Cloud edit',
	);
	const pptx = await JSZip.loadAsync(cloud.files.get('pptx').bytes);
	assert.equal(
		Object.keys(pptx.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/u.test(name)).length,
		2,
	);
	assert.match(JSON.stringify(await parseVsdx(cloud.files.get('vsdx').bytes)), /Cloud edit/u);
	await checkConcurrentMerges({ page, cloud, openCloud, edit });
	await openCloud('docx');
	await page
		.locator('.document-host:not([hidden]) docx-editor [contenteditable=true]')
		.first()
		.click();
	await page.keyboard.press('Control+A');
	await page.keyboard.type('Local version retained.');
	const remote = cloud.files.get('docx');
	remote.bytes = Buffer.from(await createDocx(['Remote version retained.']));
	remote.revision++;
	await page.locator('#save').click();
	await expect(page.locator('#onedrive-conflict')).toBeVisible();
	assert.match(JSON.stringify(await inspectDocx(remote.bytes)), /Remote version retained/u);
	await page.locator('[data-cloud-recovery=copy]').click();
	await expect(page.locator('#onedrive-conflict')).toBeHidden();
	await expect.poll(() => cloud.writes.some((id) => id.startsWith('copy-'))).toBe(true);
	await expect(page.locator('#save-state')).toHaveText('Saved to OneDrive');
	const copied = [...cloud.files.values()].find((file) => file.id.startsWith('copy-'));
	assert.ok(copied, 'conflict recovery should create a cloud copy');
	assert.match(JSON.stringify(await inspectDocx(copied.bytes)), /Local version retained/u);
	assert.match(JSON.stringify(await inspectDocx(remote.bytes)), /Remote version retained/u);
	console.log('Conflict recovery saved a separate cloud copy and preserved the remote version');
	await openCloud('xlsx');
	await edit('xlsx', 'Local Excel retained');
	const remoteExcel = cloud.files.get('xlsx');
	remoteExcel.bytes = Buffer.from(
		(
			await setXlsxCells(await createXlsx(['Sheet1']), 0, [
				{ address: 'A1', input: 'Remote Excel retained' },
			])
		).bytes,
	);
	remoteExcel.revision++;
	await page.locator('#save').click();
	await expect(page.locator('#onedrive-conflict')).toBeVisible();
	const downloaded = page.waitForEvent('download');
	await page.locator('[data-cloud-recovery=download-latest]').click();
	assert.equal(
		(await readXlsxRange(await readFile(await (await downloaded).path()), 0, 'A1')).cells[0].value,
		'Remote Excel retained',
	);
	await expect(page.locator('#onedrive-conflict')).toBeVisible();
	await page.locator('[data-cloud-recovery=latest]').click();
	await expect(page.locator('#save-state')).toHaveText('Saved to OneDrive');
	const library = await page.evaluate(async () => {
		const db = await new Promise((resolve, reject) => {
			const request = indexedDB.open('ooxml-suite-documents');
			request.onsuccess = () => resolve(request.result);
			request.onerror = () => reject(request.error);
		});
		try {
			return await new Promise((resolve, reject) => {
				const request = db.transaction('documents').objectStore('documents').getAll();
				request.onsuccess = () =>
					resolve(request.result.map((doc) => ({ ...doc, bytes: Array.from(doc.bytes) })));
				request.onerror = () => reject(request.error);
			});
		} finally {
			db.close();
		}
	});
	const backup = library.find((doc) => doc.kind === 'xlsx' && doc.name.includes('(local edits)'));
	assert.ok(backup);
	assert.equal(
		(await readXlsxRange(new Uint8Array(backup.bytes), 0, 'A1')).cells[0].value,
		'Local Excel retained',
	);
	for (const doc of library.filter((doc) => doc.cloud)) {
		assert.equal('webUrl' in doc.cloud.file, false);
		assert.equal('accessToken' in doc.cloud.file, false);
	}
	await edit('xlsx', 'Pending Excel retained');
	await page.evaluate(() => {
		globalThis.__oneDriveAuthFail = true;
	});
	await page.locator('#save').click();
	await expect(page.locator('#save-state')).toContainText('OneDrive save pending');
	await expect(page.locator(':light(#status)')).toContainText('Sign in to Microsoft again');
	await expect(page.locator(':light(#status)')).not.toContainText('Private token provider failure');
	await page.reload();
	await expect(page.locator('#save-state')).toContainText('OneDrive save pending');
	await page.locator('#save').click();
	await expect(page.locator('#save-state')).toHaveText('Saved to OneDrive');
	assert.equal(
		(await readXlsxRange(remoteExcel.bytes, 0, 'A1')).cells[0].value,
		'Pending Excel retained',
	);
	console.log(
		'Latest-version recovery keeps a local backup; failed saves survive reload and retry',
	);
	await page.evaluate(() => {
		globalThis.__oneDriveReadOnly = true;
	});
	for (const kind of ['docx', 'xlsx', 'pptx', 'vsdx']) {
		await openCloud(kind);
		await expect(page.locator('#save')).toBeDisabled();
		await expect(page.locator('#save-state')).toContainText('read-only');
		const host = page.locator('.document-host:not([hidden])');
		if (kind === 'pptx')
			await expect(host.getByRole('button', { name: 'New Slide', exact: true })).toBeDisabled();
		else
			assert.equal(
				await host
					.locator({ docx: 'docx-editor', xlsx: 'xlsx-editor', vsdx: 'visio-viewer' }[kind])
					.evaluate((editor) => editor.readOnly),
				true,
			);
	}
	console.log('Read-only permissions disable the editor and cloud Save for all four products');
	assert.deepEqual(errors, []);
	await checkCustomStorage({ browser, base, bytes: fixtures.docx });
} finally {
	await context.close();
	await browser.close();
}
