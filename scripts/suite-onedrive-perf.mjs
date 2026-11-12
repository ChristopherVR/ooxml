/**
 * Time to open a ~5 MB file from OneDrive (Graph test double) in a library that already holds
 * several large documents. Prints one line per open: node scripts/suite-onedrive-perf.mjs
 */
import { createRequire } from 'node:module';
import { interceptOneDrive } from './test-support/onedrive-http.mjs';

const require = createRequire(new URL('../e2e/docx/package.json', import.meta.url));
const { chromium, expect } = require('@playwright/test');
const JSZip = require('jszip');
const { createDocx } = await import('../src/core/dist/automation/index.mjs');

const zip = await JSZip.loadAsync(await createDocx(['Large shared document.']));
const pad = new Uint8Array(5 * 1024 * 1024);
for (let i = 0; i < pad.length; i += 65536)
	crypto.getRandomValues(pad.subarray(i, Math.min(i + 65536, pad.length)));
zip.file('customXml/pad.bin', pad, { compression: 'STORE' });
const big = await zip.generateAsync({ type: 'uint8array', compression: 'STORE' });

const base = process.env.SUITE_URL ?? 'http://127.0.0.1:8161';
const filler = Number(process.env.FILLER ?? 12);
const browser = await chromium.launch();
const context = await browser.newContext({
	viewport: { width: 1440, height: 1000 },
	serviceWorkers: 'block',
});
await context.addInitScript(() => {
	globalThis.ooxmlSkipStart = true;
	// These checks press the ribbon layout's controls; the suite's default is Simple view.
	try {
		if (!localStorage.getItem('ooxml-suite-settings'))
			localStorage.setItem('ooxml-suite-settings', JSON.stringify({ layout: 'office' }));
	} catch {}
	globalThis.ooxmlOneDriveOptions = { getAccessToken: async () => 'test-only-token' };
});
await interceptOneDrive(context, { docx: big, xlsx: big, pptx: big, vsdx: big });
const page = await context.newPage();
page.setDefaultTimeout(60000);
try {
	await page.goto(`${base}/#/`);
	await page.locator('#files').waitFor({ state: 'attached' });
	await page.waitForFunction(async () =>
		(await indexedDB.databases()).some((d) => d.name === 'ooxml-suite-documents'),
	);
	// Seed local documents (3 MB each) straight into the library database.
	await page.evaluate(async (count) => {
		const db = await new Promise((resolve, reject) => {
			const request = indexedDB.open('ooxml-suite-documents');
			request.onsuccess = () => resolve(request.result);
			request.onerror = () => reject(request.error);
		});
		await new Promise((resolve, reject) => {
			const tx = db.transaction('documents', 'readwrite');
			for (let i = 0; i < count; i++) {
				const bytes = new Uint8Array(3 * 1024 * 1024);
				tx.objectStore('documents').put({
					id: `filler-${i}`,
					name: `Filler ${i}.docx`,
					kind: 'docx',
					bytes,
					revision: 1,
					modified: Date.now(),
				});
			}
			tx.oncomplete = resolve;
			tx.onabort = () => reject(tx.error);
		});
		db.close();
	}, filler);
	await page.reload();
	const times = [];
	for (let run = 0; run < 4; run++) {
		if (!(await page.locator('#open-onedrive').isVisible()))
			await page.locator('#rail [data-home]').click();
		await page.locator('#open-onedrive').click();
		const link = page.locator('#onedrive-link, office-ui-cloud-files .link input').first();
		await link.fill('https://tenant.sharepoint.com/links/docx');
		const started = Date.now();
		await link.press('Enter');
		await expect(page.locator('#save-state')).toContainText('OneDrive');
		times.push(Date.now() - started);
		await page
			.locator('#tabs [data-close]')
			.first()
			.click()
			.catch(() => {});
	}
	// The lookup an open performs: every stored document (old) against the bytes-free index (new).
	const lookup = await page.evaluate(async () => {
		const db = await new Promise((resolve, reject) => {
			const request = indexedDB.open('ooxml-suite-documents');
			request.onsuccess = () => resolve(request.result);
			request.onerror = () => reject(request.error);
		});
		const time = async (name) => {
			if (!db.objectStoreNames.contains(name)) return undefined;
			const started = performance.now();
			await new Promise((resolve, reject) => {
				const request = db.transaction(name).objectStore(name).getAll();
				request.onsuccess = resolve;
				request.onerror = () => reject(request.error);
			});
			return Math.round(performance.now() - started);
		};
		const result = { documents: await time('documents'), cloud: await time('cloud') };
		db.close();
		return result;
	});
	console.log(
		'lookup (ms): getAll(documents)=' + lookup.documents + ', getAll(cloud index)=' + lookup.cloud,
	);
	console.log(`filler=${filler}x3MB, open 5MB via link (ms): ${times.join(', ')}`);
} finally {
	await browser.close();
}
