import { readFile } from 'node:fs/promises';

import { expect, test, type Page } from '@playwright/test';
import JSZip from 'jszip';

import { loadDeck, SAMPLE_DECK, slideElements, slideStage } from './support/deck';

test.use({ viewport: { width: 375, height: 812 }, hasTouch: true, isMobile: true });

const picker = (page: Page, name: string) =>
	page.locator('office-ui-dialog[open]').filter({
		has: page.getByRole('heading', { name, exact: true }),
	});

test.beforeEach(async ({ page }) => {
	await page.addInitScript(() => {
		const devices = navigator.mediaDevices;
		Object.defineProperty(devices, 'getDisplayMedia', {
			configurable: true,
			value: () => {
				const canvas = document.createElement('canvas');
				canvas.width = 320;
				canvas.height = 180;
				const context = canvas.getContext('2d')!;
				let frame = 0;
				const tick = () => {
					context.fillStyle = frame++ % 2 ? '#0451a5' : '#00aa77';
					context.fillRect(0, 0, canvas.width, canvas.height);
				};
				tick();
				const stream = canvas.captureStream(20);
				const timer = setInterval(() => {
					if (stream.getVideoTracks()[0]?.readyState === 'ended') clearInterval(timer);
					else tick();
				}, 50);
				return Promise.resolve(stream);
			},
		});
	});
	await loadDeck(page);
});

test('captures a screen clipping as an image on the slide', async ({ page }) => {
	const before = await slideElements(page).count();
	await page
		.getByRole('navigation', { name: 'Editor actions' })
		.getByRole('button', { name: 'Insert' })
		.tap();
	await page
		.locator('pptx-ui-mobile-insert')
		.getByRole('button', { name: 'Screenshot', exact: true })
		.tap();
	await page
		.locator('pptx-ui-mobile-insert')
		.getByRole('button', { name: 'Screen Clipping', exact: true })
		.tap();
	const dialog = picker(page, 'Screen Clipping');
	await dialog.getByRole('spinbutton', { name: /^(W|Width)$/ }).fill('100');
	await dialog.getByRole('button', { name: 'Insert', exact: true }).tap();
	await expect(slideElements(page)).toHaveCount(before + 1);
	await expect(slideStage(page).getByRole('img', { name: 'Screenshot' })).toBeVisible();
});

test('embeds a local screen recording as slide video', async ({ page }) => {
	const before = await slideElements(page).count();
	await page
		.getByRole('navigation', { name: 'Editor actions' })
		.getByRole('button', { name: 'Insert' })
		.tap();
	await page
		.locator('pptx-ui-mobile-insert')
		.getByRole('button', { name: 'Screen Recording', exact: true })
		.tap();
	const dialog = picker(page, 'Screen Recording');
	await dialog.getByRole('button', { name: 'Record', exact: true }).tap();
	await dialog.getByRole('button', { name: 'Pause', exact: true }).tap();
	await dialog.getByRole('button', { name: 'Resume', exact: true }).tap();
	await dialog.getByRole('button', { name: 'Stop', exact: true }).tap();
	await expect(slideElements(page)).toHaveCount(before + 1);
	const pending = page.waitForEvent('download');
	await page
		.getByRole('toolbar', { name: 'Toolbar', exact: true })
		.getByRole('button', { name: 'Save', exact: true })
		.tap();
	const archive = await JSZip.loadAsync(await readFile(await (await pending).path()));
	expect(Object.keys(archive.files)).toEqual(
		expect.arrayContaining([expect.stringMatching(/^ppt[/]media[/]video\d+\.(mp4|webm)$/u)]),
	);
});

test('creates and downloads a separate Photo Album', async ({ page }) => {
	const count = await page.evaluate(() => window.__pptxViewer.getSlides().length);
	await page
		.getByRole('navigation', { name: 'Editor actions' })
		.getByRole('button', { name: 'Insert' })
		.tap();
	await page
		.locator('pptx-ui-mobile-insert')
		.getByRole('button', { name: 'Photo Album', exact: true })
		.tap();
	const dialog = picker(page, 'Photo Album');
	await dialog.getByRole('textbox', { name: 'Album title' }).fill('Field trip');
	const image = Buffer.from(
		await page.evaluate(() => {
			const canvas = document.createElement('canvas');
			canvas.width = 50;
			canvas.height = 40;
			const context = canvas.getContext('2d')!;
			context.fillStyle = '#0451a5';
			context.fillRect(0, 0, 50, 40);
			return canvas.toDataURL('image/png').split(',')[1]!;
		}),
		'base64',
	);
	await dialog.locator('input[type=file]').setInputFiles([
		{ name: 'Terrace.png', mimeType: 'image/png', buffer: image },
		{ name: 'Waterfall.png', mimeType: 'image/png', buffer: image },
	]);
	await dialog.getByRole('combobox', { name: 'Pictures per slide' }).click();
	await dialog.getByRole('option', { name: '2', exact: true }).click();
	await dialog.getByRole('checkbox', { name: 'Captions below pictures' }).check();
	const pending = page.waitForEvent('download');
	await dialog.getByRole('button', { name: 'Create Album', exact: true }).tap();
	const download = await pending;
	expect(download.suggestedFilename()).toBe('Field trip.pptx');
	const archive = await JSZip.loadAsync(await readFile(await download.path()));
	expect(await archive.file('ppt/slides/slide2.xml')?.async('string')).toContain('Waterfall');
	expect(await page.evaluate(() => window.__pptxViewer.getSlides().length)).toBe(count);
});

test('reuses only the selected slide from a local presentation', async ({ page }) => {
	const count = await page.evaluate(() => window.__pptxViewer.getSlides().length);
	await page
		.getByRole('navigation', { name: 'Editor actions' })
		.getByRole('button', { name: 'Insert' })
		.tap();
	await page
		.locator('pptx-ui-mobile-insert')
		.getByRole('button', { name: 'Reuse Slides', exact: true })
		.tap();
	const dialog = picker(page, 'Reuse Slides');
	await dialog.locator('input[type=file]').setInputFiles(SAMPLE_DECK);
	await dialog.getByRole('checkbox', { name: /^Slide 2\b/ }).check();
	await dialog.getByRole('button', { name: 'Insert', exact: true }).tap();
	await expect
		.poll(() => page.evaluate(() => window.__pptxViewer.getSlides().length))
		.toBe(count + 1);
	await expect(slideStage(page).getByText('Agenda', { exact: true })).toBeVisible();
});
