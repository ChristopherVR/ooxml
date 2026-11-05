import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { interceptOneDrive } from './onedrive-http.mjs';

const require = createRequire(new URL('../../e2e/docx/package.json', import.meta.url));
const { expect } = require('@playwright/test');

const row = (page, name) =>
	page.locator('#cloud-files li[data-entry]').filter({ hasText: name }).first();
/** The connected documents of the library as `itemId:documentId`, from its bytes-free index. */
const connected = (page) =>
	page.evaluate(async () => {
		const db = await new Promise((resolve, reject) => {
			const request = indexedDB.open('ooxml-suite-documents');
			request.onsuccess = () => resolve(request.result);
			request.onerror = () => reject(request.error);
		});
		try {
			return await new Promise((resolve, reject) => {
				const request = db.transaction('cloud').objectStore('cloud').getAll();
				request.onsuccess = () =>
					resolve(request.result.map((doc) => `${doc.cloud.file.itemId}:${doc.id}`));
				request.onerror = () => reject(request.error);
			});
		} finally {
			db.close();
		}
	});

/**
 * The OneDrive workspace view on the host-signed-in page of the main check: Home entry, Recent,
 * My files, a folder, search, a non-Office file, then each format opened from the browser into its
 * editor, edited and saved to the same item. Reopening reuses the library document.
 */
export async function checkBrowse({ page, cloud, edit, nativeSave }) {
	await page.evaluate(() => {
		globalThis.__oneDriveReadOnly = false;
	});
	await page.locator('#rail [data-home]').click();
	await expect(page.locator('#open-onedrive')).toHaveText('Browse OneDrive');
	await page.locator('#open-onedrive').click();
	await expect(page.locator('#cloud-view')).toBeVisible();
	await expect(page.locator('#rail [data-cloud]')).toHaveAttribute('aria-current', 'true');
	// No host sign-in hook: the host is already signed in and the account reads generically.
	await expect(page.locator('#cloud-files .account')).toHaveText('Microsoft account');
	await expect(page.locator('#cloud-files [data-view=recent]')).toBeVisible();
	await page.locator('#cloud-files [data-view=files]').click();
	await expect(row(page, 'Projects')).toBeVisible();
	await expect(row(page, 'Shared.docx')).toBeVisible();
	await expect(row(page, 'Notes.txt')).toHaveAttribute('aria-disabled', 'true');
	assert.ok(cloud.requests.some((r) => r.startsWith('GET /me/drive/root/children')));
	// Listings select only the fields they need.
	assert.ok(
		cloud.requests.filter((r) => r.includes('/children')).every((r) => r.includes('select')),
	);
	// Search finds files in folders.
	await page.locator('#cloud-files office-ui-search').evaluate((search) => {
		search.value = 'Shared.pptx';
		search.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
	});
	await expect(row(page, 'Shared.pptx')).toBeVisible();
	for (const kind of ['docx', 'xlsx', 'pptx', 'vsdx']) {
		await page.locator('#rail [data-cloud]').click();
		await page.locator('#cloud-files [data-view=files]').click();
		if (kind === 'pptx') await expect(row(page, 'Notes.txt')).toBeVisible();
		if (kind === 'pptx' || kind === 'vsdx') {
			// The browser keeps its folder between visits: enter Projects only from the root.
			await expect(page.locator('#cloud-files [data-crumb]').first()).toBeVisible();
			if ((await page.locator('#cloud-files [data-crumb]').count()) === 1) {
				await expect(row(page, 'Notes.txt')).toBeVisible();
				await row(page, 'Projects').click();
			}
			await expect(page.locator('#cloud-files [data-crumb]').last()).toHaveText('Projects');
			assert.ok(
				cloud.requests.some((r) => r.startsWith('GET /drives/test-drive/items/projects/children')),
			);
		}
		const writesBefore = cloud.writes.filter((id) => id === kind).length;
		await row(page, `Shared.${kind}`).click();
		await expect(page.locator('#save-state')).toContainText('OneDrive');
		await expect(page.locator('#cloud-view')).toBeHidden();
		// The same cloud file is one library document however it was opened.
		assert.equal(
			(await connected(page)).filter((entry) => entry.startsWith(`${kind}:`)).length,
			1,
			`one library document for ${kind}`,
		);
		await edit(kind, `Browsed ${kind}`);
		await nativeSave(kind);
		await expect.poll(() => cloud.writes.filter((id) => id === kind).length).toBe(writesBefore + 1);
		await expect(page.locator('#save-state')).toContainText('Saved to OneDrive');
	}
	// Reopening from the browser reuses the existing document.
	const before = await connected(page);
	await page.locator('#rail [data-cloud]').click();
	await page.locator('#cloud-files [data-crumb="0"]').click();
	await row(page, 'Shared.docx').click();
	await expect(page.locator('#cloud-view')).toBeHidden();
	assert.deepEqual(await connected(page), before);
	console.log(
		'Browse: Home entry, Recent, folders, search and all four formats opened from OneDrive',
	);
}

/** A host with a sign-in hook starts signed out; signing in runs the hook and lists the files. */
export async function checkHostSignIn({ browser, base, fixtures }) {
	const context = await browser.newContext({ serviceWorkers: 'block' });
	await context.addInitScript(() => {
		globalThis.ooxmlSkipStart = true;
		globalThis.__signIns = 0;
		globalThis.__signOuts = 0;
		globalThis.ooxmlOneDriveOptions = {
			getAccessToken: async () => 'test-only-token',
			signIn: async () => {
				globalThis.__signIns++;
				return { name: 'Ada Lovelace', email: 'ada@example.test' };
			},
			signOut: async () => {
				globalThis.__signOuts++;
			},
		};
	});
	await interceptOneDrive(context, fixtures);
	const page = await context.newPage();
	const errors = [];
	page.on('pageerror', (error) => errors.push(error.message));
	try {
		await page.goto(`${base}/#/`);
		await page.locator('#rail [data-cloud]').click();
		await expect(page.locator('#cloud-files .signin')).toBeVisible();
		assert.equal(await page.evaluate(() => globalThis.__signIns), 0);
		await page.locator('#cloud-files .signin .primary').click();
		await expect(page.locator('#cloud-files .account')).toHaveText('Ada Lovelace');
		assert.equal(await page.evaluate(() => globalThis.__signIns), 1);
		await page.locator('#cloud-files [data-view=files]').click();
		await expect(row(page, 'Shared.docx')).toBeVisible();
		await page.locator('#cloud-files [data-action=sign-out]').click();
		await expect(page.locator('#cloud-files .signin')).toBeVisible();
		assert.equal(await page.evaluate(() => globalThis.__signOuts), 1);
		assert.deepEqual(errors, []);
		console.log('Host sign-in hook: signed out, signed in, listed, signed out');
	} finally {
		await context.close();
	}
}

/** Without any host options the view keeps the honest "not connected" message. */
export async function checkUnconnected({ browser, base }) {
	const context = await browser.newContext({ serviceWorkers: 'block' });
	await context.addInitScript(() => {
		globalThis.ooxmlSkipStart = true;
	});
	const page = await context.newPage();
	try {
		await page.goto(`${base}/#/`);
		await page.locator('#rail [data-cloud]').click();
		await expect(page.locator('#onedrive-open-form')).toContainText(
			'Microsoft sign-in has not been connected by your host',
		);
		console.log('Unconnected host: honest message, no browser');
	} finally {
		await context.close();
	}
}
