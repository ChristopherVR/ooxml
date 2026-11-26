/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { savePptxViaBackstage } from './save-pptx';
import { elementsOfType, fixture, loadDeck, selectElement } from './support/deck';
import { downloadBytes } from './support/exports';
import { readZipPartText } from './support/pptx-xml';
import { chooseSelectValue } from './support/select-control';

test.use({ viewport: { width: 2048, height: 1080 } });
const control = (page: Page, id: string) => page.locator(`[data-ribbon-control="${id}"]:visible`);
const tab = (page: Page, id: string) => page.locator(`[data-ribbon-contextual-tab="${id}"]`);

test('video poster, crop and playback edits save to the slide', async ({ page }) => {
	await loadDeck(page, fixture('media-bookmark-trigger-editable.pptx'));
	await selectElement(page, elementsOfType(page, 'media player').first());
	await expect(tab(page, 'videoFormat')).toBeVisible();
	await tab(page, 'videoFormat').click();
	const video = page.locator('[data-pptx-viewport] video').first();
	await expect
		.poll(() => video.evaluate((clip: HTMLVideoElement) => clip.videoWidth))
		.toBeGreaterThan(0);
	const originalPoster = await video.getAttribute('poster');
	await control(page, 'videoFormat.poster.current').getByRole('button').click();
	await expect(video).not.toHaveAttribute('poster', originalPoster ?? '');
	await expect(page.locator('[data-media-ribbon-dialog="current"]')).toHaveCount(0);
	await control(page, 'videoFormat.poster.file').getByRole('button').click();
	const poster = page.locator('[data-media-ribbon-dialog="poster"]');
	const png = await page.evaluate(() => {
		const canvas = document.createElement('canvas');
		canvas.width = 12;
		canvas.height = 8;
		const context = canvas.getContext('2d')!;
		context.fillStyle = '#bb3355';
		context.fillRect(0, 0, 12, 8);
		return canvas.toDataURL().split(',')[1]!;
	});
	await poster.locator('input[type="file"]').setInputFiles({
		name: 'poster.png',
		mimeType: 'image/png',
		buffer: Buffer.from(png, 'base64'),
	});
	await poster.getByRole('button', { name: 'Apply' }).click();
	await expect(poster).toHaveCount(0);
	const importedPoster = await video.getAttribute('poster');
	await control(page, 'videoFormat.poster.reset').getByRole('button').click();
	await expect(video).not.toHaveAttribute('poster', importedPoster ?? '');
	await expect(page.locator('[data-media-ribbon-dialog="reset"]')).toHaveCount(0);
	await control(page, 'videoFormat.pictureStyles.shape').getByRole('button').click();
	const ellipse = page.locator('[data-ribbon-gallery-popup] [data-gallery-item="ellipse"]:visible');
	await expect(ellipse.locator('svg path')).toBeVisible();
	expect(await video.evaluate((clip) => getComputedStyle(clip).clipPath)).toBe('none');
	await ellipse.click();
	await expect(video).toHaveCSS('clip-path', /^ellipse\(/u);
	await control(page, 'videoFormat.arrange.rotateRight').getByRole('button').click();
	await control(page, 'videoFormat.size.crop').getByRole('button').click();
	const crop = page.locator('[data-media-ribbon-dialog="crop"]');
	await crop.getByLabel('Crop Left %').fill('16');
	await crop.getByRole('button', { name: 'Apply' }).click();
	await expect(crop).toHaveCount(0);

	await tab(page, 'playback').click();
	const remove = control(page, 'playback.bookmarks.remove').getByRole('button');
	await remove.click();
	const dialog = page.locator('[data-media-ribbon-dialog="remove"]');
	await expect(dialog.locator('pptx-ui-select option')).toHaveCount(2);
	await expect(dialog.locator('pptx-ui-select option').first()).toHaveText(
		/^BM1 \(\d+:\d{2}\.\d\)$/u,
	);
	await dialog.getByRole('button', { name: 'Apply' }).click();
	await expect(dialog).toHaveCount(0);
	await control(page, 'playback.bookmarks.add').getByRole('button').click();

	await control(page, 'playback.editing.trim').getByRole('button').click();
	const trim = page.locator('[data-media-ribbon-dialog="trim"]');
	await trim.getByLabel('Trim Start (s)').fill('0.15');
	const trimEnd = trim.getByLabel('Trim End (s)');
	await trimEnd.fill('0.1');
	const applyTrim = trim.getByRole('button', { name: 'Apply' });
	await expect(applyTrim).toBeDisabled();
	await trimEnd.fill('1.5');
	await applyTrim.click();
	await expect(trim).toHaveCount(0);

	const fade = control(page, 'playback.editing.fadeIn').getByRole('spinbutton');
	await fade.fill('0.25');
	await fade.press('Tab');
	const volume = control(page, 'playback.options.volume').getByRole('combobox');
	await chooseSelectValue(page, volume, 'medium');
	await chooseSelectValue(page, volume, 'mute');
	await expect(video).toHaveJSProperty('volume', 0);
	const loop = control(page, 'playback.options.loop').getByRole('checkbox');
	await loop.click();
	await expect(loop).toBeChecked();
	await chooseSelectValue(
		page,
		control(page, 'playback.options.start').getByRole('combobox'),
		'sequence',
	);
	const xml = await readZipPartText(
		await downloadBytes(await savePptxViaBackstage(page)),
		'ppt/slides/slide1.xml',
	);
	expect(xml).toContain('l="16000"');
	expect(xml).toContain('prst="ellipse"');
	expect(xml).toContain('rot="5400000"');
	expect(xml).toContain('<p14:fade in="250"');
	expect(xml).toContain('<p14:trim st="150" end="1500"');
	expect(xml).not.toContain('<p14:bmk name="BM1"');
	expect(xml).toContain('<p14:bmk name="BM2"');
	expect(xml).toMatch(/<p14:bmk name="Bookmark \d+" time="\d+"/u);
	expect(xml).toContain('vol="50000"');
	expect(xml).toContain('mute="1"');
	expect(xml).toContain('repeatCount="indefinite"');
	expect(xml).toContain('cmd="playFrom(0.0)"');
});

test('a media dialog returns focus to its ribbon control inside a shadow root', async ({
	page,
}) => {
	await loadDeck(page, fixture('media-bookmark-trigger-editable.pptx'));
	await selectElement(page, elementsOfType(page, 'media player').first());
	await tab(page, 'playback').click();
	const remove = control(page, 'playback.bookmarks.remove').getByRole('button');
	await expect(remove).toBeEnabled();
	await remove.evaluate((button) => {
		const command = (button.getRootNode() as ShadowRoot).host;
		const strip = command.parentElement!;
		const wrapper = document.createElement('div');
		strip.before(wrapper);
		wrapper.attachShadow({ mode: 'open' }).append(strip);
	});
	await remove.focus();
	await remove.press('Enter');
	const dialog = page.locator('[data-media-ribbon-dialog="remove"]');
	await expect(dialog).toBeVisible();
	await dialog.getByRole('button', { name: 'Cancel' }).click();
	await expect(dialog).toHaveCount(0);
	await expect(remove).toBeFocused();
});

test('audio shows its own Format tab and media-specific playback options', async ({ page }) => {
	await loadDeck(page, fixture('audio-embed.pptx'));
	await selectElement(page, elementsOfType(page, 'media player').first());
	await expect(tab(page, 'audioFormat')).toBeVisible();
	await expect(tab(page, 'videoFormat')).toHaveCount(0);
	await tab(page, 'audioFormat').click();
	await expect(control(page, 'audioFormat.size.width').getByRole('spinbutton')).toBeEnabled();
	await control(page, 'audioFormat.pictureStyles.shape').getByRole('button').click();
	await page.locator('[data-ribbon-gallery-popup] [data-gallery-item="roundRect"]:visible').click();
	await tab(page, 'playback').click();
	const fullScreen = control(page, 'playback.options.fullScreen').getByRole('checkbox');
	await expect(fullScreen).toBeDisabled();
	await expect(fullScreen).toHaveAccessibleDescription('This command is for video.');
	const across = control(page, 'playback.options.across').getByRole('checkbox');
	await expect(across).toBeEnabled();
	if (!(await across.isChecked())) await across.click();
	await chooseSelectValue(
		page,
		control(page, 'playback.options.start').getByRole('combobox'),
		'auto',
	);
	await chooseSelectValue(
		page,
		control(page, 'playback.options.start').getByRole('combobox'),
		'click',
	);
	await expect(across).toBeChecked();
	const xml = await readZipPartText(
		await downloadBytes(await savePptxViaBackstage(page)),
		'ppt/slides/slide1.xml',
	);
	expect(xml).toContain('numSld="999"');
	expect(xml).toContain('prst="roundRect"');
});

test('SmartArt Format applies galleries to the canvas node selection', async ({ page }) => {
	await loadDeck(page, fixture('smartart-build-reveal.pptx'));
	const node = (name: string) =>
		page.locator('[data-pptx-viewport] [data-smartart-node-id]').filter({ hasText: name }).first();
	const alpha = node('Alpha');
	const beta = node('Beta');
	await selectElement(page, alpha);
	await expect(tab(page, 'smartArtDesign')).toBeVisible();
	await tab(page, 'smartArtFormat').click();
	await beta.click({ modifiers: ['Control'] });
	await expect(page.locator('[data-pptx-viewport] [data-pptx-smartart-selected]')).toHaveCount(2);
	const bold = control(page, 'smartArtFormat.wordArtStyles.bold').getByRole('checkbox');
	await bold.click();
	await expect(bold).toBeChecked();
	const wordArt = control(page, 'smartArtFormat.wordArtStyles.gallery');
	await wordArt.locator('[data-gallery-item="wordArt-1"]').first().click();
	const shapeStyles = control(page, 'smartArtFormat.shapeStyles.gallery');
	await shapeStyles.locator('[data-gallery-item="theme-0-1"]').first().click();
	await control(page, 'smartArtFormat.shapes.shape').getByRole('button').click();
	await page.locator('[data-ribbon-gallery-popup] [data-gallery-item="hexagon"]:visible').click();
	await expect(
		page
			.locator('[data-ribbon-group="smartArtFormat.arrange"]')
			.getByRole('button', { name: /^bring forward$/iu })
			.first(),
	).toBeVisible();
	const xml = await readZipPartText(
		await downloadBytes(await savePptxViaBackstage(page)),
		'ppt/diagrams/drawing1.xml',
	);
	expect(xml.match(/prst="hexagon"/gu)).toHaveLength(2);
});
