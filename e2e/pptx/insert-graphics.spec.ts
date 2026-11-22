/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { readFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

import { savePptxViaBackstage } from './save-pptx';
import { chooseCommand, openMenuOn } from './support/context-menu';
import {
	loadDeck,
	ribbonTab,
	selectElement,
	slideElements,
	slideStage,
	thumbnail,
} from './support/deck';
import { openFormatSection } from './support/format-pane';

test.use({ viewport: { width: 1440, height: 900 } });

test('inserts icons, symbols, editable WordArt and a Slide Zoom, then reopens the deck', async ({
	page,
}, testInfo) => {
	await loadDeck(page);
	await ribbonTab(page, 'Insert').click();
	const ribbon = page.locator('pptx-ui-ribbon-insert');
	const stage = page.locator('[data-pptx-viewport]');
	const originals = await slideElements(page).count();
	const open = (name: string) => ribbon.getByRole('button', { name, exact: true }).click();
	const picker = ribbon.locator('office-ui-dialog');
	const submit = async (heading: string) => {
		await expect(page.getByRole('dialog', { name: heading })).toBeVisible();
		await picker.getByRole('button', { name: 'Insert', exact: true }).click();
	};

	await open('Icons');
	await page.getByRole('searchbox', { name: 'Search icons' }).fill('light bulb');
	await page.getByRole('button', { name: 'Light bulb', exact: true }).click();
	await submit('Icons');
	await expect(slideElements(page)).toHaveCount(originals + 1);

	await open('Symbol');
	await page.getByRole('textbox', { name: 'Character code' }).fill('U+2603');
	await page.getByRole('textbox', { name: 'Character code' }).press('Enter');
	await expect(stage.getByText('☃', { exact: true })).toBeVisible();

	await open('WordArt');
	await picker.getByRole('textbox', { name: 'WordArt text' }).fill('Quarterly highlights');
	await picker.locator('.pptx-wordart-grid button').nth(3).click();
	await submit('WordArt');
	await expect(stage.getByText('Quarterly highlights', { exact: true })).toBeVisible();

	await ribbonTab(page, 'Home').click();
	await stage.getByText('Quarterly highlights', { exact: true }).dblclick();
	const editor = page.locator('[data-inline-editor]');
	await expect(editor).toBeFocused();
	await editor.press('Home');
	await editor.press('ArrowRight');
	await ribbonTab(page, 'Insert').click();
	await open('Symbol');
	await page.getByRole('textbox', { name: 'Character code' }).fill('U+2605');
	await page.getByRole('textbox', { name: 'Character code' }).press('Enter');
	await expect(editor).toHaveText('Q★uarterly highlights');
	await editor.press('Escape');
	await expect(slideElements(page)).toHaveCount(originals + 3);

	await open('Zoom');
	await ribbon.locator('[data-insert-item="slideZoom"]').click();
	await page.getByRole('radio', { name: /^Slide 2\b/ }).check();
	await submit('Slide Zoom');
	await expect(slideElements(page)).toHaveCount(originals + 4);

	const path = testInfo.outputPath('insert-graphics.pptx');
	await (await savePptxViaBackstage(page)).saveAs(path);
	const saved = await JSZip.loadAsync(await readFile(path));
	const slide = await saved.file('ppt/slides/slide1.xml')?.async('string');
	expect(slide).toContain('asvg:svgBlip');
	expect(slide).toContain('pslz:sldZm');
	await loadDeck(page, path);
	await expect(slideElements(page)).toHaveCount(originals + 4);
	await expect(stage.getByText('☃', { exact: true })).toBeVisible();
	await expect(stage.getByText('Q★uarterly highlights', { exact: true })).toBeVisible();
	const zoomPreview = stage.locator('[data-zoom-type="slide"][data-zoom-target="1"] img');
	await expect(zoomPreview).toBeVisible();
	await expect
		.poll(() =>
			zoomPreview.evaluate((image) => (image instanceof HTMLImageElement ? image.naturalWidth : 0)),
		)
		.toBeGreaterThan(0);
});

test('inserts and reopens Section and Summary Zoom previews for existing sections', async ({
	page,
}, testInfo) => {
	await loadDeck(page);
	for (const number of [2, 4]) {
		await openMenuOn(page, thumbnail(page, number));
		await chooseCommand(page, 'add section');
	}
	await thumbnail(page, 1).click();
	await ribbonTab(page, 'Insert').click();
	const ribbon = page.locator('pptx-ui-ribbon-insert');
	const picker = ribbon.locator('office-ui-dialog');
	const stage = page.locator('[data-pptx-viewport]');
	const selectZoom = async (type: string) => {
		await ribbon.getByRole('button', { name: 'Zoom', exact: true }).click();
		await ribbon.locator(`[data-insert-item="${type}"]`).click();
	};
	const insert = () => picker.getByRole('button', { name: 'Insert', exact: true }).click();
	await selectZoom('sectionZoom');
	await picker.getByRole('radio').first().check();
	await insert();
	await expect(stage.locator('[data-zoom-type="section"][data-zoom-target="1"] img')).toBeVisible();
	await selectZoom('summaryZoom');
	await expect(picker.getByRole('checkbox', { name: /^Slide/, checked: true })).toHaveCount(2);
	await insert();
	await expect(stage.locator('[data-zoom-type="summary"] img')).toHaveCount(2);

	const path = testInfo.outputPath('insert-section-zooms.pptx');
	await (await savePptxViaBackstage(page)).saveAs(path);
	const reopened = await page.context().newPage();
	await loadDeck(reopened, path);
	await thumbnail(reopened, 1).click();
	const reopenedStage = slideStage(reopened);
	await expect(
		reopenedStage.locator('[data-zoom-type="section"][data-zoom-target="2"] img'),
	).toBeVisible();
	await thumbnail(reopened, 2).click();
	await expect(reopenedStage.locator('[data-zoom-type="summary"] img')).toHaveCount(2);
});

test('creates sections from selected slides in a deck without any sections', async ({ page }) => {
	await loadDeck(page);
	await ribbonTab(page, 'Insert').click();
	const ribbon = page.locator('pptx-ui-ribbon-insert');
	await ribbon.getByRole('button', { name: 'Zoom', exact: true }).click();
	await ribbon.locator('[data-insert-item="summaryZoom"]').click();
	const dialog = ribbon.locator('office-ui-dialog[open]');
	await expect(dialog.getByRole('checkbox', { name: /^Slide \d+\b/, checked: true })).toHaveCount(
		0,
	);
	await dialog.getByRole('checkbox', { name: /^Slide 1\b/ }).check();
	await dialog.getByRole('checkbox', { name: /^Slide 4\b/ }).check();
	await dialog.getByRole('button', { name: 'Insert', exact: true }).click();
	const stage = slideStage(page);
	await expect(stage.locator('[data-zoom-type="summary"] img')).toHaveCount(2);
	await expect(thumbnail(page, 8)).toBeVisible();
	await page.getByRole('button', { name: 'Undo', exact: true }).first().click();
	await expect(thumbnail(page, 8)).toHaveCount(0);
	await expect(stage.locator('[data-zoom-type="summary"]')).toHaveCount(0);
	await page.getByRole('button', { name: 'Redo', exact: true }).first().click();
	await expect(thumbnail(page, 8)).toBeVisible();
	await expect(stage.locator('[data-zoom-type="summary"] img')).toHaveCount(2);
});

async function embeddedZoomPng(path: string): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(await readFile(path));
	const xml = await zip.file('ppt/slides/slide1.xml')?.async('string');
	const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')?.async('string');
	const id = xml?.match(/<pslz:zmPr\b[^>]*>\s*<p166:blipFill>\s*<a:blip\s+r:embed="([^"]+)"/u)?.[1];
	const relationship = rels
		?.match(/<Relationship\b[^>]*>/gu)
		?.find((entry) => id && entry.includes(`Id="${id}"`));
	const image = relationship?.match(/Target="\.\.\/(media\/[^"]+\.png)"/u)?.[1];
	if (!image) throw new Error('Slide Zoom has no native PNG relationship');
	const bytes = await zip.file(`ppt/${image}`)?.async('uint8array');
	if (!bytes?.length) throw new Error('Slide Zoom PNG is missing from the package');
	return bytes;
}

test('Slide Zoom saves destination edits and lets the Format pane replace and reset the image', async ({
	page,
}, testInfo) => {
	await loadDeck(page);
	await ribbonTab(page, 'Insert').click();
	const ribbon = page.locator('pptx-ui-ribbon-insert');
	await ribbon.getByRole('button', { name: 'Zoom', exact: true }).click();
	await ribbon.locator('[data-insert-item="slideZoom"]').click();
	const dialog = ribbon.locator('office-ui-dialog[open]');
	await dialog.getByRole('radio', { name: /^Slide 2\b/ }).check();
	await dialog.getByRole('button', { name: 'Insert', exact: true }).click();
	await expect(slideStage(page).locator('[data-zoom-type="slide"] img')).toBeVisible();

	const before = testInfo.outputPath('original-zoom.pptx');
	await (await savePptxViaBackstage(page)).saveAs(before);
	const original = await embeddedZoomPng(before);
	await thumbnail(page, 2).click();
	await ribbonTab(page, 'Insert').click();
	await ribbon.getByRole('button', { name: 'WordArt', exact: true }).click();
	await dialog.getByRole('textbox', { name: 'WordArt text' }).fill('New Zoom destination heading');
	await dialog.locator('.pptx-wordart-grid button').nth(4).click();
	await dialog.getByRole('button', { name: 'Insert', exact: true }).click();
	await expect(slideStage(page).getByText('New Zoom destination heading')).toBeVisible();

	const after = testInfo.outputPath('refreshed-zoom.pptx');
	await (await savePptxViaBackstage(page)).saveAs(after);
	const refreshed = await embeddedZoomPng(after);
	expect(refreshed).not.toEqual(original);

	await thumbnail(page, 1).click();
	const zoom = slideStage(page).locator('[data-zoom-type="slide"]');
	await selectElement(page, zoom);
	const sizePosition = await openFormatSection(page, 'sizePosition');
	const cover = sizePosition.locator('pptx-ui-zoom-cover');
	const choose = cover.getByRole('button', { name: 'Change Image', exact: true });
	const reset = cover.getByRole('button', { name: 'Reset Image', exact: true });
	await expect(reset).toBeDisabled();
	const coverBase64 = await page.evaluate(() => {
		const canvas = document.createElement('canvas');
		canvas.width = 24;
		canvas.height = 16;
		const context = canvas.getContext('2d')!;
		context.fillStyle = '#102b8b';
		context.fillRect(0, 0, canvas.width, canvas.height);
		return canvas.toDataURL().split(',')[1]!;
	});
	const choosing = page.waitForEvent('filechooser');
	await choose.click();
	await (
		await choosing
	).setFiles({
		name: 'cover.png',
		mimeType: 'image/png',
		buffer: Buffer.from(coverBase64, 'base64'),
	});
	await expect(reset).toBeEnabled();
	const covered = testInfo.outputPath('covered-zoom.pptx');
	await (await savePptxViaBackstage(page)).saveAs(covered);
	expect(await embeddedZoomPng(covered)).not.toEqual(refreshed);

	await reset.click();
	await expect(reset).toBeDisabled();
	const restored = testInfo.outputPath('restored-zoom.pptx');
	await (await savePptxViaBackstage(page)).saveAs(restored);
	expect(await embeddedZoomPng(restored)).toEqual(refreshed);
});
