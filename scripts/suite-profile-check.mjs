import { createRequire } from 'node:module';
import { resolve, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';
const root = resolve(fileURLToPath(new URL('..', import.meta.url)));
const require = createRequire(join(root, 'e2e/docx/package.json'));
const { chromium, expect } = require('@playwright/test');
const { createDocx } = await import('../src/core/dist/automation/index.mjs');
const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({ viewport: { width: 1440, height: 960 } });
// Skip the start chooser; these checks drive the whole suite.
await page.addInitScript(() => (globalThis.ooxmlSkipStart = true));
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
const base = process.env.SUITE_URL ?? 'http://127.0.0.1:8130';
const shot = (name) =>
	page.screenshot({
		path: join(process.env.TEMP ?? root, `ooxml-profile-${name}.png`),
		animations: 'disabled',
	});
try {
	await page.goto(base);
	await page.locator('#profile-toggle').click();
	await page.locator('#account-menu').getByRole('button', { name: 'View my profile' }).click();
	await page.locator('[name=name]').fill('Alex Morgan');
	await page.locator('[name=email]').fill('alex@example.com');
	await page.locator('[name=organization]').fill('Northwind Studio');
	await page.locator('[name=title]').fill('Design lead');
	await page.locator('[name=photo]').setInputFiles({
		name: 'avatar.png',
		mimeType: 'image/png',
		buffer: Buffer.from(
			'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a5ioAAAAASUVORK5CYII=',
			'base64',
		),
	});
	await expect(page.locator('#photo-preview img')).toBeVisible();
	await page.locator('[data-remove-photo]').click();
	await expect(page.locator('#photo-preview img')).toHaveCount(0);
	await page.getByRole('button', { name: 'Save profile', exact: true }).click();
	assert.equal(await page.locator('#workspace-name').innerText(), 'Northwind Studio');
	await page.reload();
	await page.locator('#profile-toggle').click();
	await page.getByRole('heading', { name: 'Alex Morgan', exact: true }).waitFor();
	await shot('account');
	await page.keyboard.press('Escape');
	await expect(page.locator('#profile-toggle')).toHaveAttribute('aria-expanded', 'false');
	await page.locator('#file-input').setInputFiles({
		name: 'Design brief.docx',
		mimeType: 'application/octet-stream',
		buffer: Buffer.from(await createDocx(['Design brief', 'Initial project scope.'])),
	});
	await page.locator('docx-editor [contenteditable=true]').first().waitFor();
	await expect(page.locator('docx-editor [contenteditable=true]').first()).toContainText(
		'Initial project scope.',
	);
	await page.locator('docx-editor [contenteditable=true]').first().click();
	await page.keyboard.press('Control+End');
	await page.keyboard.type(' Added before switching profiles.');
	await page.locator('#rail [data-teams]').click();
	await page.waitForFunction(
		() => document.querySelector('teams-app')?.client?.getState().user.name === 'Alex Morgan',
	);
	await page.locator('#rail [data-home]').click();
	await page.locator('[data-favourite]').click();
	await page.locator('#rail [data-library=favourites]').click();
	await expect(page.locator('#rail [aria-current=true]')).toHaveCount(1);
	assert.equal(await page.locator('#files tr').count(), 1);
	await page.locator('#rail [data-home]').click();
	await shot('home');
	await page.locator('.workspace-switch').click();
	await page.locator('[data-profile-add]').click();
	await page.locator('[name=name]').fill('Personal');
	await page.getByRole('button', { name: 'Create profile', exact: true }).click();
	await page.waitForFunction(
		() =>
			document.querySelector('#profile-toggle')?.getAttribute('aria-label') ===
			'Account manager for Personal',
	);
	assert.equal(await page.locator('#files tr').count(), 0);
	await page.locator('#rail [data-teams]').click();
	await page.waitForFunction(
		() => document.querySelector('teams-app')?.client?.getState().user.name === 'Personal',
	);
	await page.locator('.workspace-switch').click();
	await page.locator('[data-switch-profile=default]').click();
	await page.waitForFunction(
		() =>
			document.querySelector('#profile-toggle')?.getAttribute('aria-label') ===
			'Account manager for Alex Morgan',
	);
	await page.locator('#files tr').waitFor();
	assert.equal(await page.locator('#files tr').count(), 1);
	await page.locator('#files [data-open]').click();
	await page.locator('docx-editor [contenteditable=true]').first().waitFor();
	await expect(page.locator('docx-editor [contenteditable=true]').first()).toContainText(
		'Added before switching profiles',
	);
	await page.locator('#rail [data-home]').click();
	await page.locator('#rail [data-library=favourites]').click();
	await expect(page.locator('#rail [aria-current=true]')).toHaveCount(1);
	assert.equal(await page.locator('#files tr').count(), 1);
	await page.locator('#rail [data-home]').click();
	await page.locator('#theme-picker').click();
	await page.locator('[name=mode][value=dark]').check();
	await page.locator('[name=accent][value=purple]').check();
	await page.getByRole('button', { name: 'Done', exact: true }).click();
	await shot('dark');
	await page.setViewportSize({ width: 390, height: 844 });
	await shot('mobile');
	assert(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
	await page.locator('#mobile-search-toggle').click();
	await page.locator('#search').fill('Design');
	assert.equal(await page.locator('#files tr').count(), 1);
	await page.locator('#search').fill('');
	await page.locator('#mobile-search-toggle').click();
	await page.locator('#profile-toggle').click();
	await shot('mobile-account');
	const box = await page.locator('#account-menu').boundingBox();
	assert(box.x >= 0 && box.x + box.width <= 390);
	await page.locator('#account-menu').getByRole('button', { name: 'View my profile' }).click();
	await shot('mobile-form');
	const formBox = await page.locator('.profile-form').boundingBox();
	assert(formBox.x >= 0 && formBox.x + formBox.width <= 390);
	assert.equal(
		await page
			.locator('.profile-fields input')
			.first()
			.evaluate((e) => getComputedStyle(e).fontSize),
		'16px',
	);
	await page.locator('.form-footer [data-cancel]').click();
	await page.locator('#launcher-toggle').click();
	await page.locator('#app-launcher [data-app=xlsx]').click();
	await page.getByRole('heading', { name: 'Excel', exact: true }).waitFor();
	assert.deepEqual(errors, []);
	console.log(
		'Profile editing/persistence, Teams identity, save-before-switch, library isolation, favourites, launcher, themes and mobile checks passed.',
	);
} finally {
	await browser.close();
}
