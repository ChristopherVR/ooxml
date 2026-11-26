/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { savePptxViaBackstage } from './save-pptx';
import {
	elementsOfType,
	elementWithText,
	fixture,
	loadDeck,
	openRibbonTab,
	selectElement,
} from './support/deck';
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
	await dialog.getByRole('button', { name: 'Apply' }).click();
	await expect(dialog).toHaveCount(0);

	const fade = control(page, 'playback.editing.fadeIn').getByRole('spinbutton');
	await fade.fill('0.25');
	await fade.press('Tab');
	await control(page, 'playback.options.volume').getByRole('spinbutton').fill('50');
	await control(page, 'playback.options.volume').getByRole('spinbutton').press('Tab');
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
	expect(xml).toContain('rot="5400000"');
	expect(xml).toContain('<p14:fade in="250"');
	expect(xml).not.toContain('<p14:bmk name="BM1"');
	expect(xml).toContain('<p14:bmk name="BM2"');
	expect(xml).toContain('vol="50000"');
	expect(xml).toContain('repeatCount="indefinite"');
	expect(xml).toContain('cmd="playFrom(0.0)"');
});

test('audio shows its own Format tab and media-specific playback options', async ({ page }) => {
	await loadDeck(page, fixture('audio-embed.pptx'));
	await selectElement(page, elementsOfType(page, 'media player').first());
	await expect(tab(page, 'audioFormat')).toBeVisible();
	await expect(tab(page, 'videoFormat')).toHaveCount(0);
	await tab(page, 'audioFormat').click();
	await expect(control(page, 'audioFormat.size.width').getByRole('spinbutton')).toBeEnabled();
	await tab(page, 'playback').click();
	await expect(control(page, 'playback.options.fullScreen').getByRole('checkbox')).toBeDisabled();
	await expect(control(page, 'playback.options.across').getByRole('checkbox')).toBeEnabled();
});

test('SmartArt Format edits the node selected in the ribbon', async ({ page }) => {
	await loadDeck(page, fixture('smartart-build-reveal.pptx'));
	await selectElement(page, elementWithText(page, 'Alpha'));
	await expect(tab(page, 'smartArtDesign')).toBeVisible();
	await tab(page, 'smartArtFormat').click();
	const nodes = control(page, 'smartArtFormat.shapes.node').getByRole('combobox');
	const chooseNode = async (label: string) => {
		await nodes.click();
		await page.getByRole('option', { name: label, exact: true }).click();
	};
	await chooseNode('Beta');
	const bold = control(page, 'smartArtFormat.wordArtStyles.bold').getByRole('checkbox');
	await bold.click();
	await expect(bold).toBeChecked();
	await chooseNode('Alpha');
	await expect(bold).not.toBeChecked();
	await chooseNode('Beta');
	await expect(bold).toBeChecked();
	await openRibbonTab(page, 'Home');
	await expect(nodes).not.toBeVisible();
});
