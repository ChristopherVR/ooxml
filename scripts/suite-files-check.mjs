import { createRequire } from 'node:module';
import { join } from 'node:path';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, unlink, rmdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
const require = createRequire(new URL('../e2e/docx/package.json', import.meta.url));
const { chromium, expect } = require('@playwright/test');
const { createDocx } = await import('../src/core/dist/automation/index.mjs');
const browser = await chromium.launch({ channel: 'chromium' });
const context = await browser.newContext({ viewport: { width: 1440, height: 960 } });
// Skip the start chooser; these checks drive the whole suite.
await context.addInitScript(() => sessionStorage.setItem('ooxml-start', 'office'));
const page = await context.newPage();
page.setDefaultTimeout(15000);
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
let fixturePath;
try {
	await page.goto(process.env.SUITE_URL || 'http://127.0.0.1:8133');
	await page.locator('#profile-toggle').click();
	await expect(page.locator('#account-menu')).toBeVisible();
	await expect(page.locator('#account-menu [data-switch-profile]')).toHaveCount(0);
	await page.locator('.workspace-switch').click();
	await expect(page.locator('#workspace-menu')).toBeVisible();
	await expect(page.locator('#account-menu')).toBeHidden();
	await expect(page.locator('#workspace-menu [data-switch-profile]')).toHaveCount(1);
	await page.locator('#storage-toggle').click();
	await expect(page.locator('#storage-menu')).toBeVisible();
	await expect(page.locator('#workspace-menu')).toBeHidden();
	await page.keyboard.press('Escape');
	const bytes = Array.from(await createDocx(['Disk source, keep this original.']));
	await page.evaluate(async (bytes) => {
		const root = await navigator.storage.getDirectory();
		const folder = await root.getDirectoryHandle('Fixture folder', { create: true });
		for (const name of ['brief.docx', 'budget.docx']) {
			const handle = await folder.getFileHandle(name, { create: true });
			const writer = await handle.createWritable();
			await writer.write(new Uint8Array(bytes));
			await writer.close();
		}
		window.showDirectoryPicker = async () => folder;
	}, bytes);
	await page.locator('#storage-toggle').click();
	await page.locator('#storage-menu [data-connect-folder]').click();
	await page.keyboard.press('Escape');
	await page.locator('#search').fill('brief');
	await expect(page.locator('#disk-search [data-disk-open]')).toHaveCount(1);
	await expect(page.locator('#disk-search')).toContainText('brief.docx');
	await page.locator('#search').fill('bud');
	await page.locator('#search').fill('brief');
	await page.locator('#search').fill('budget');
	await expect(page.locator('#disk-search [data-disk-open]')).toHaveCount(1);
	await expect(page.locator('#disk-search')).toContainText('budget.docx');
	await expect(page.locator('#disk-search')).not.toContainText('brief.docx');
	await page.locator('#disk-search [data-disk-open]').click();
	await expect(page.locator('docx-editor')).toBeVisible();
	const id = await page.locator('.file-tab [data-open]').getAttribute('data-open');
	await page.locator('.file-tab [data-open]').click({ button: 'right' });
	await page.getByRole('menuitem', { name: 'File location', exact: true }).click();
	await expect(page.locator('#dialog')).toContainText('Fixture folder/budget.docx');
	await page.locator('#dialog-close').click();
	const editor = page.locator('docx-editor [contenteditable=true]').first();
	await editor.click();
	await page.keyboard.press('Control+End');
	await page.keyboard.type(' Unsaved tab edit.');
	await page.locator('.file-tab [data-open]').click({ button: 'middle' });
	await expect(page.locator('.file-tab')).toHaveCount(0);
	await page.locator('#search').fill('');
	await page.locator(`#files [data-open="${id}"]`).click();
	await expect(page.locator('docx-editor [contenteditable=true]').first()).toContainText(
		'Unsaved tab edit.',
	);
	for (let n = 0; n < 2; n++) {
		await page.locator('#rail [data-home]').click();
		await page.locator('[data-create=docx]').click();
		await expect(page.locator('.file-tab')).toHaveCount(n + 2);
	}
	await page.locator('.file-tab [data-open]').first().click({ button: 'right' });
	await page.getByRole('menuitem', { name: 'Close tabs to the right' }).click();
	await expect(page.locator('.file-tab')).toHaveCount(1);
	await page.locator('#rail [data-home]').click();
	await page.locator(`#files [data-file-menu="${id}"]`).click();
	await page.getByRole('menuitem', { name: 'Remove from workspace…' }).click();
	await page.locator('[data-confirm-remove]').click();
	await expect(page.locator(`#files [data-open="${id}"]`)).toHaveCount(0);
	await expect(page.locator('.file-tab')).toHaveCount(0);
	const original = await page.evaluate(async () => {
		const root = await navigator.storage.getDirectory();
		return Array.from(
			new Uint8Array(
				await (
					await (await root.getDirectoryHandle('Fixture folder')).getFileHandle('budget.docx')
				)
					.getFile()
					.then((f) => f.arrayBuffer()),
			),
		);
	});
	assert.deepEqual(original, bytes);
	await page.locator('.file-undo').click();
	await expect(page.locator(`#files [data-open="${id}"]`)).toHaveCount(1);
	// Chromium crashes when using an OPFS test handle deserialized after navigation.
	// Verify persisted metadata here; real disk permission prompts remain platform-owned.
	const savedFolders = await page.evaluate(
		() =>
			new Promise((resolve) => {
				const request = indexedDB.open('ooxml-suite-documents-locations');
				request.onsuccess = () => {
					const tx = request.result.transaction('folders');
					const read = tx.objectStore('folders').getAllKeys();
					read.onsuccess = () => resolve(read.result);
				};
			}),
	);
	assert.equal(savedFolders.length, 1);
	await page.locator('#search').fill('budget');
	await expect(page.locator('#disk-search [data-disk-open]')).toHaveCount(1);
	await page.locator('#storage-toggle').click();
	await page.locator('[data-disconnect]').click();
	await page.keyboard.press('Escape');
	await expect(page.locator('#disk-search')).toContainText('Connect a folder');

	// Test the browser fallback with a real directory selection and no persistent handle API.
	fixturePath = await mkdtemp(join(tmpdir(), 'ooxml-folder-check-'));
	await writeFile(join(fixturePath, 'snapshot.docx'), new Uint8Array(bytes));
	await page.evaluate(() => {
		window.showDirectoryPicker = undefined;
	});
	await page.locator('#storage-toggle').click();
	const picker = page.waitForEvent('filechooser');
	await page.locator('#storage-menu [data-connect-folder]').click();
	await (await picker).setFiles(fixturePath);
	await page.keyboard.press('Escape');
	await page.locator('#search').fill('snapshot');
	await expect(page.locator('#disk-search [data-disk-open]')).toHaveCount(1);
	await page.reload();
	await page.locator('#search').fill('snapshot');
	await expect(page.locator('#disk-search')).toContainText('Connect a folder');
	await page.locator('#search').fill('');
	await expect(page.locator('#library-title')).toHaveText('Your documents');
	await page.setViewportSize({ width: 390, height: 844 });
	await page.locator('[data-storage]').first().click();
	const box = await page.locator('#storage-menu').boundingBox();
	assert(box.x >= 0 && box.x + box.width <= 390);
	await page.screenshot({
		path: join(process.env.TEMP, 'ooxml-storage-phone.png'),
		animations: 'disabled',
	});
	assert.deepEqual(errors, []);
	console.log(
		'Distinct menus, connected folder search, snapshot fallback, stale query protection, original-file preservation, remove/undo, tab close/save and context menus passed.',
	);
} finally {
	await browser.close();
	if (fixturePath) {
		await unlink(join(fixturePath, 'snapshot.docx'));
		await rmdir(fixturePath);
	}
}
