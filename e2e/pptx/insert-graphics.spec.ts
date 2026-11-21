/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { readFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

import { savePptxViaBackstage } from './save-pptx';
import { chooseCommand, openMenuOn } from './support/context-menu';
import { loadDeck, ribbonTab, slideElements, slideStage, thumbnail } from './support/deck';

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
	await expect(dialog.getByRole('checkbox', { checked: true })).toHaveCount(0);
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
