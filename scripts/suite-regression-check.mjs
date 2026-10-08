import { createRequire } from 'node:module';
import { join } from 'node:path';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
const require = createRequire(new URL('../e2e/docx/package.json', import.meta.url));
const { chromium, expect } = require('@playwright/test');
const base = process.env.SUITE_URL || 'http://127.0.0.1:8133';
const browser = await chromium.launch();
const context = await browser.newContext({ viewport: { width: 1440, height: 960 } });
// Skip the first-visit start chooser; these checks drive the whole suite.
await context.addInitScript(() => localStorage.setItem('ooxml-start-app', 'office'));
const page = await context.newPage();
page.setDefaultTimeout(20000);
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
try {
	await page.goto(base);
	const sidebar = await page.locator('.workspace-sidebar').boundingBox();
	await page.locator('#rail [data-teams]').click();
	await page.locator('office-ui-chat-composer textarea').waitFor();
	assert.equal((await page.locator('.workspace-sidebar').boundingBox()).width, sidebar.width);
	await expect(page.locator('teams-app teams-profile-menu')).toBeHidden();
	const composer = page.locator('office-ui-chat-composer');
	const area = composer.locator('textarea');
	const payloads = [
		{
			name: 'Brief.pdf',
			mimeType: 'application/pdf',
			buffer: Buffer.from('%PDF-1.4\nTest attachment\n%%EOF'),
		},
		{
			name: 'photo.png',
			mimeType: 'image/png',
			buffer: Buffer.from(
				'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a5ioAAAAASUVORK5CYII=',
				'base64',
			),
		},
		{
			name: 'archive.zip',
			mimeType: 'application/zip',
			buffer: Buffer.from([80, 75, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
		},
	];
	await composer.locator('input[type=file]').setInputFiles(payloads);
	await expect(area).toBeFocused();
	await area.fill('General file attachments');
	await area.press('Enter');
	const list = page.locator('office-ui-chat-list');
	await expect(list.getByRole('button', { name: /Brief.pdf/ })).toBeVisible();
	await expect(area).toBeFocused();
	await page.reload();
	await page.locator('office-ui-chat-composer textarea').waitFor();
	for (const file of payloads) {
		await list.getByRole('button', { name: new RegExp(file.name.replace('.', '\\.')) }).click();
		const pending = page.waitForEvent('download');
		await page.locator('#dialog').getByRole('button', { name: 'Download file' }).click();
		const download = await pending;
		assert.equal(download.suggestedFilename(), file.name);
		assert.deepEqual(await readFile(await download.path()), file.buffer);
		await page.locator('#dialog-close').click();
	}
	const longName = 'A-very-long-attachment-name-'.repeat(6) + '.txt';
	await page.setViewportSize({ width: 390, height: 844 });
	await composer
		.locator('input[type=file]')
		.setInputFiles({ name: longName, mimeType: 'text/plain', buffer: Buffer.from('hi') });
	await expect(area).toBeFocused();
	const bounds = await composer.boundingBox();
	assert(bounds.x + bounds.width <= 391);
	const over = await composer.evaluate((el) => ({
		scroll: el.shadowRoot.querySelector('.box').scrollWidth,
		width: el.shadowRoot.querySelector('.box').clientWidth,
	}));
	assert(over.scroll <= over.width + 1, JSON.stringify(over));
	await composer.getByRole('button', { name: `Remove ${longName}` }).click();
	await expect(area).toBeFocused();
	await area.fill('one two');
	await area.evaluate((el) => el.setSelectionRange(4, 4));
	await composer.getByRole('button', { name: 'Emoji', exact: true }).click();
	await composer.getByRole('menuitem').first().click();
	assert.match(await area.inputValue(), /^one .+two$/);
	await expect(area).toBeFocused();
	await page.screenshot({ path: join(process.env.TEMP, 'suite-teams-phone.png') });
	console.log(
		'Teams: PDF/PNG/ZIP persist and download byte-for-byte; focus, emoji caret and phone overflow passed.',
	);
	await page.setViewportSize({ width: 1440, height: 960 });
	await page.locator('#rail [data-home]').click();
	const start = Date.now();
	await page.locator('[data-create=pptx]').click();
	await page.locator('.pptxv-stage[role=region]').waitFor();
	console.log('Local PowerPoint opening ms:', Date.now() - start);
	assert.equal((await page.locator('.workspace-sidebar').boundingBox()).width, sidebar.width);
	for (const mode of ['light', 'dark']) {
		await page.locator('#theme-picker').click();
		await page.locator(`input[name=mode][value=${mode}]`).check();
		await page.locator('input[name=accent][value=green]').check();
		await page.locator('#dialog').getByRole('button', { name: 'Done' }).click();
		for (let i = 0; i < 2; i++) {
			await page.locator('.pptxv button[title="Settings & Shortcuts"]').click();
			const dialog = page.locator('.pptxv-options-dialog');
			await expect(dialog).toBeVisible();
			const palette = await dialog.evaluate((el) => ({
				bg: getComputedStyle(el).backgroundColor,
				token: getComputedStyle(el).getPropertyValue('--pptx-primary').trim(),
			}));
			assert.equal(palette.bg, mode === 'dark' ? 'rgb(32, 34, 40)' : 'rgb(255, 255, 255)');
			assert.equal(palette.token, mode === 'dark' ? '#54b054' : '#107c41');
			await dialog.getByRole('button', { name: 'OK', exact: true }).click();
			await expect(dialog).toHaveCount(0);
			assert.equal(await page.locator('.pptxv-parity-backdrop').count(), 0);
		}
		const overflow = await page.locator('.pptxv').evaluate((el) =>
			[...el.querySelectorAll('[class*=inspector] input,[class*=inspector] select')]
				.filter((e) => e.getBoundingClientRect().width > 0)
				.filter((e) => {
					const a = e.getBoundingClientRect(),
						b = e.closest('[class*=inspector]').getBoundingClientRect();
					return a.right > b.right + 2;
				})
				.map((e) => e.outerHTML),
		);
		assert.deepEqual(overflow, []);
	}
	await page.screenshot({ path: join(process.env.TEMP, 'suite-pptx-dark-fixed.png') });
	await page.locator('.pptxv button[title="Settings & Shortcuts"]').click();
	await page.setViewportSize({ width: 390, height: 844 });
	const box = await page.locator('.pptxv-options-dialog').boundingBox();
	assert(box.x >= 0 && box.x + box.width <= 391 && box.y + box.height <= 844);
	await page.screenshot({ path: join(process.env.TEMP, 'suite-pptx-settings-phone.png') });
	await page.keyboard.press('Escape');
	console.log(
		'PowerPoint: light/dark/accent, repeat cog close, properties control bounds and phone settings passed.',
	);
	await page.evaluate(() => navigator.serviceWorker.ready);
	await page.waitForFunction(() => !!navigator.serviceWorker.controller, {}, { timeout: 60000 });
	const ids = new Set();
	for (const path of [
		'/',
		'/apps/word/',
		'/apps/excel/',
		'/apps/powerpoint/',
		'/apps/visio/',
		'/apps/teams/',
	]) {
		const response = await context.request.get(base + path + 'manifest.webmanifest');
		assert(response.ok());
		const manifest = await response.json();
		ids.add(new URL(manifest.id, base + path).href);
	}
	assert.equal(ids.size, 6);
	await context.setOffline(true);
	for (const [slug, kind] of [
		['word', 'docx'],
		['excel', 'xlsx'],
		['powerpoint', 'pptx'],
		['visio', 'vsdx'],
		['teams', 'teams'],
	]) {
		await page.goto(base + '/apps/' + slug + '/');
		await expect(page.locator('body')).toHaveAttribute('data-product', kind);
		await expect(page.locator('.workspace-sidebar')).toBeHidden();
		if (kind === 'teams')
			await expect(page.locator('office-ui-chat-composer textarea')).toBeVisible();
		else {
			await page.locator(`[data-create=${kind}]`).click();
			const editor = { pptx: '.pptxv', vsdx: 'visio-viewer' }[kind] ?? `${kind}-editor`;
			await expect(page.locator(editor)).toBeVisible();
		}
	}
	await page.goto(base + '/');
	await page.locator('[data-create=docx]').click();
	await expect(page.locator('docx-editor')).toBeVisible();
	await context.setOffline(false);
	// A first visit asks what to open, remembers it, and ?suite returns to the whole suite.
	const fresh = await browser.newContext();
	const visitor = await fresh.newPage();
	await visitor.goto(base + '/');
	await visitor.locator('#dialog [data-start=docx]').click();
	await visitor.waitForURL(/\/apps\/word\/$/);
	await visitor.goto(base + '/');
	await visitor.waitForURL(/\/apps\/word\/$/);
	await visitor.locator('#launcher-toggle').click();
	await visitor.locator('#app-launcher a', { hasText: 'OOXML Office' }).click();
	await expect(visitor.locator('[data-create=docx]')).toBeVisible();
	await expect(visitor.locator('#dialog')).not.toBeVisible();
	await visitor.goto(base + '/');
	await expect(visitor).toHaveURL(base + '/');
	await fresh.close();
	assert.deepEqual(errors, []);
	console.log(
		'Start chooser remembers the app. PWA: six distinct manifests; all standalone pages and suite reopen offline; Word/Excel/PowerPoint/Visio create offline.',
	);
} finally {
	await browser.close();
}
