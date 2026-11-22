import { readFile } from 'node:fs/promises';

import { expect, test, type Page } from '@playwright/test';
import JSZip from 'jszip';

import { loadDeck, slideElements, slideStage } from './support/deck';

test.use({ viewport: { width: 375, height: 812 }, hasTouch: true, isMobile: true });

const tiles = (page: Page) => page.locator('pptx-ui-mobile-insert');
const insertSheet = (page: Page) =>
	page.getByRole('navigation', { name: 'Editor actions' }).getByRole('button', { name: 'Insert' });

async function openPicker(page: Page, name: string) {
	await tiles(page).getByRole('button', { name, exact: true }).tap();
	await expect(page.getByRole('dialog', { name, exact: true })).toBeVisible();
	return page.locator('office-ui-dialog[open]').filter({
		has: page.getByRole('heading', { name, exact: true }),
	});
}

test.beforeEach(async ({ page }) => {
	await loadDeck(page);
});

test('phone Insert can open, cancel and commit WordArt', async ({ page }) => {
	await insertSheet(page).tap();
	const originalCount = await slideElements(page).count();
	const picker = await openPicker(page, 'WordArt');
	await page.keyboard.press('Escape');
	await expect(picker).toHaveCount(0);
	await expect(
		page.locator('pptx-ui-mobile-insert').getByRole('button', { name: 'WordArt' }),
	).toBeFocused();

	const reopened = await openPicker(page, 'WordArt');
	await reopened.getByRole('textbox', { name: 'WordArt text' }).fill('Design review');
	await reopened.locator('.pptx-wordart-grid button').first().tap();
	await reopened.getByRole('button', { name: 'Insert', exact: true }).tap();

	await expect(reopened).toHaveCount(0);
	await expect(slideElements(page)).toHaveCount(originalCount + 1);
	await expect(tiles(page)).toBeHidden();
	await expect(slideStage(page).getByText('Design review', { exact: true })).toBeVisible();
});

test('phone Insert keeps selected icons while searching and inserts them together', async ({
	page,
}) => {
	await insertSheet(page).tap();
	const originalCount = await slideElements(page).count();
	const picker = await openPicker(page, 'Icons');
	await picker.getByRole('searchbox', { name: 'Search icons' }).fill('Light bulb');
	await picker.getByRole('button', { name: 'Light bulb', exact: true }).tap();
	await picker.getByRole('searchbox', { name: 'Search icons' }).fill('Check mark');
	await picker.getByRole('button', { name: 'Check mark', exact: true }).tap();
	await picker.getByRole('button', { name: 'Insert', exact: true }).tap();

	await expect(picker).toHaveCount(0);
	await expect(slideElements(page)).toHaveCount(originalCount + 2);
	await expect(tiles(page)).toBeHidden();
	for (const name of ['Light bulb', 'Check mark']) {
		await expect(slideStage(page).getByRole('img', { name, exact: true })).toBeVisible();
	}
});

test('phone Insert can add a character from the Symbol picker', async ({ page }) => {
	await insertSheet(page).tap();
	const originalCount = await slideElements(page).count();
	const picker = await openPicker(page, 'Symbol');
	const code = picker.getByRole('textbox', { name: 'Character code' });
	await code.fill('U+2603');
	await code.press('Enter');

	await expect(picker).toHaveCount(0);
	await expect(slideElements(page)).toHaveCount(originalCount + 1);
	await expect(tiles(page)).toBeHidden();
	await expect(slideStage(page).getByText('☃', { exact: true })).toBeVisible();
});

test('phone Zoom inserts Summary, Section and Slide Zoom', async ({ page }) => {
	await insertSheet(page).tap();
	await tiles(page).getByRole('button', { name: 'Zoom', exact: true }).tap();
	const chooser = await openPicker(page, 'Summary Zoom');
	await chooser.getByRole('checkbox', { name: /^Slide 2\b/ }).tap();
	await chooser.getByRole('checkbox', { name: /^Slide 4\b/ }).tap();
	await chooser.getByRole('button', { name: 'Insert', exact: true }).tap();

	await expect(tiles(page)).toBeHidden();
	await expect(slideStage(page).locator('[data-zoom-type="summary"] img')).toHaveCount(2);

	for (const type of ['Section', 'Slide']) {
		await insertSheet(page).tap();
		await tiles(page).getByRole('button', { name: 'Zoom', exact: true }).tap();
		const picker = await openPicker(page, `${type} Zoom`);
		await picker.getByRole('radio').first().tap();
		await picker.getByRole('button', { name: 'Insert', exact: true }).tap();

		await expect(tiles(page)).toBeHidden();
		const preview = slideStage(page).locator(`[data-zoom-type="${type.toLowerCase()}"] img`);
		await expect(preview).toBeVisible();
	}
	const zoom = slideStage(page).locator('[data-zoom-type="slide"]');
	const preview = zoom.locator('img');
	await expect(preview).toHaveAttribute('src', /^data:image\/svg\+xml/);
	const initialImage = await preview.getAttribute('src');
	const destination = Number(await zoom.getAttribute('data-zoom-target'));
	await page.evaluate(async (index) => {
		const api = window.__pptxViewer;
		const slide = api.getSlides()[index];
		const shape = slide?.elements.find((element) => element.type === 'text');
		if (!slide || !shape) throw new Error('Missing text on the Zoom destination');
		await api.updateElements([
			{ slideId: slide.id, elementId: shape.id, patch: { x: shape.x + 120 } },
		]);
	}, destination);
	await expect.poll(() => preview.getAttribute('src')).not.toBe(initialImage);
});

test('phone Shapes inserts the chosen preset and Undo, Redo and Save act on it', async ({
	page,
}) => {
	const before = await slideElements(page).count();
	await insertSheet(page).tap();
	await tiles(page).getByRole('button', { name: 'Shapes', exact: true }).tap();
	await tiles(page).getByRole('button', { name: 'Circle', exact: true }).tap();
	await expect(tiles(page)).toBeHidden();
	await expect(slideElements(page)).toHaveCount(before + 1);

	const toolbar = page.getByRole('toolbar', { name: 'Toolbar', exact: true });
	await toolbar.getByRole('button', { name: 'Undo', exact: true }).tap();
	await expect(slideElements(page)).toHaveCount(before);
	await toolbar.getByRole('button', { name: 'Redo', exact: true }).tap();
	await expect(slideElements(page)).toHaveCount(before + 1);

	const pending = page.waitForEvent('download');
	await toolbar.getByRole('button', { name: 'Save', exact: true }).tap();
	const file = await (await pending).path();
	const deck = await JSZip.loadAsync(await readFile(file));
	const slide = await deck.file('ppt/slides/slide1.xml')?.async('string');
	expect(slide?.match(/<p:sp[\s>][\s\S]*?<\/p:sp>/g)?.at(-1)).toContain('prst="ellipse"');
});

test('phone Insert adds the chosen table size and chart type', async ({ page }) => {
	await insertSheet(page).tap();
	await tiles(page).getByRole('button', { name: 'Table', exact: true }).tap();
	await tiles(page).getByRole('grid').getByRole('row').nth(1).getByRole('gridcell').nth(2).tap();
	await expect(tiles(page)).toBeHidden();
	const rows = slideStage(page).locator('table tr');
	await expect(rows).toHaveCount(2);
	await expect(rows.first().locator('td')).toHaveCount(3);

	await insertSheet(page).tap();
	await tiles(page).getByRole('button', { name: 'Chart', exact: true }).tap();
	await tiles(page).getByRole('button', { name: 'Pie', exact: true }).tap();
	await expect(tiles(page)).toBeHidden();
	await expect(slideStage(page).locator('path[data-chart-part="dataPoint"]').first()).toBeVisible();
});

for (const tile of ['Pictures', 'Video']) {
	test(`phone Insert > ${tile} opens the device file picker`, async ({ page }) => {
		await insertSheet(page).tap();
		const [chooser] = await Promise.all([
			page.waitForEvent('filechooser'),
			tiles(page).getByRole('button', { name: tile, exact: true }).tap(),
		]);
		expect(chooser.isMultiple()).toBe(false);
		await expect(tiles(page)).toBeHidden();
	});
}
