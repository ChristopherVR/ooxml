/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { savePptxViaBackstage } from './save-pptx';
import { goToSlide } from './support/context-menu';
import { loadDeck, ribbonTab, slideStage } from './support/deck';
import { downloadBytes } from './support/exports';
import { readZipPartText } from './support/pptx-xml';

/**
 * Round 6 Design tab: Background Styles (apply to all / selected slides, tile menu, hover preview
 * that reverts and never records an undo step), Effects (every built-in scheme, applied mark) and
 * the honest Variants heading. One framework-neutral spec, run on every binding; it reads only
 * `data-ribbon-*`, `data-gallery-*` hooks, the slide stage's computed background and accessible
 * names.
 */
test.use({ viewport: { width: 1440, height: 900 } });

async function open(page: Page): Promise<void> {
	await loadDeck(page);
	await ribbonTab(page, 'Design').click();
	await page.waitForTimeout(200);
}

const stageBackground = (page: Page): Promise<string> =>
	slideStage(page).evaluate((node) => getComputedStyle(node).backgroundColor);

const undo = (page: Page): Locator =>
	page.getByRole('button', { name: 'Undo', exact: true }).first();

async function openGallery(page: Page, control: string): Promise<Locator> {
	const root = page.locator(`[data-ribbon-control="design.variants.${control}"]`);
	await root.locator('[data-ribbon-gallery]').first().click();
	return root;
}

/** The Effects tiles marked applied, read with the gallery open, which is then closed. */
async function appliedEffects(page: Page): Promise<(string | null)[]> {
	const root = await openGallery(page, 'effects');
	const tiles = root.locator('[data-gallery-item]');
	await expect(tiles.first()).toBeVisible();
	const applied = await root
		.locator('[data-gallery-item][aria-pressed="true"]')
		.evaluateAll((nodes) => nodes.map((node) => node.getAttribute('data-gallery-item')));
	await page.keyboard.press('Escape');
	await expect(tiles.first()).toBeHidden();
	return applied;
}

/**
 * Save the deck through File > Save and read its theme part's format scheme name; the Design tab
 * is put back on screen.
 */
async function savedEffects(page: Page): Promise<string> {
	const bytes = await downloadBytes(await savePptxViaBackstage(page));
	const design = ribbonTab(page, 'Design');
	if ((await design.getAttribute('aria-selected')) !== 'true') {
		await design.click();
	}
	const xml = await readZipPartText(bytes, 'ppt/theme/theme1.xml');
	return /<a:fmtScheme\b[^>]*\bname="([^"]*)"/u.exec(xml)?.[1] ?? 'missing';
}

test.describe('Design > Background Styles', () => {
	test('lists the 12 styles with a tile menu and a Format Background footer', async ({ page }) => {
		await open(page);
		const root = await openGallery(page, 'backgroundStyles');
		await expect(root.locator('[data-gallery-item]')).toHaveCount(12);
		await expect(root.locator('[data-gallery-footer="formatBackground"]')).toHaveText(
			'Format Background...',
		);
		await root.locator('[data-gallery-item="2"]').click({ button: 'right' });
		await expect(root.getByRole('menuitem')).toHaveText([
			'Apply to All Slides',
			'Apply to Selected Slides',
		]);
	});

	test('hovering shows the style, leaving puts the slide back, and no undo step is recorded', async ({
		page,
	}) => {
		await open(page);
		const original = await stageBackground(page);
		const root = await openGallery(page, 'backgroundStyles');
		const dark = root.locator('[data-gallery-item="4"]');
		await dark.hover();
		await expect.poll(() => stageBackground(page)).not.toBe(original);
		await page.mouse.move(700, 600);
		await expect.poll(() => stageBackground(page)).toBe(original);
		await expect(undo(page)).toBeDisabled();
	});

	test('a pick applies to every slide as one undo step', async ({ page }) => {
		await open(page);
		const original = await stageBackground(page);
		const root = await openGallery(page, 'backgroundStyles');
		await root.locator('[data-gallery-item="4"]').click();
		await expect.poll(() => stageBackground(page)).not.toBe(original);
		const applied = await stageBackground(page);
		await expect(undo(page)).toBeEnabled();
		await undo(page).click();
		await expect.poll(() => stageBackground(page)).toBe(original);
		expect(applied).not.toBe(original);
	});

	test('the tile menu applies to the selected slide only', async ({ page }) => {
		await open(page);
		const original = await stageBackground(page);
		const root = await openGallery(page, 'backgroundStyles');
		await root.locator('[data-gallery-item="4"]').click({ button: 'right' });
		await root.getByRole('menuitem', { name: 'Apply to Selected Slides' }).click();
		await expect.poll(() => stageBackground(page)).not.toBe(original);
		// The next slide keeps the deck's own background (reached through the thumbnail rail).
		await goToSlide(page, 2);
		await expect.poll(() => stageBackground(page)).toBe(original);
		await goToSlide(page, 1);
		await expect.poll(() => stageBackground(page)).not.toBe(original);
		// One undo step puts the first slide back.
		await undo(page).click();
		await expect.poll(() => stageBackground(page)).toBe(original);
	});
});

test.describe('Design > Effects', () => {
	test('lists every built-in scheme and marks the one a pick applies', async ({ page }) => {
		await open(page);
		const root = await openGallery(page, 'effects');
		const tiles = root.locator('[data-gallery-item]');
		await expect(tiles).toHaveCount(14);
		await expect(root.locator('[data-gallery-item="glossy"]')).toHaveAttribute(
			'title',
			'Effects: Glossy',
		);
		await root.locator('[data-gallery-item="glossy"]').click();
		const reopened = await openGallery(page, 'effects');
		await expect(reopened.locator('[data-gallery-item="glossy"]')).toHaveAttribute(
			'aria-pressed',
			'true',
		);
	});

	test('Undo takes a pick back on the tiles and in the saved file, Redo reapplies it', async ({
		page,
	}) => {
		await open(page);
		const redo = page.getByRole('button', { name: 'Redo', exact: true }).first();
		// The sample deck's own scheme ("Office Theme") is none of the built-in tiles.
		expect(await appliedEffects(page)).toStrictEqual([]);
		const root = await openGallery(page, 'effects');
		await root.locator('[data-gallery-item="glossy"]').click();
		expect(await appliedEffects(page)).toStrictEqual(['glossy']);
		await expect(undo(page)).toBeEnabled();
		await undo(page).click();
		// Undo used to leave the format scheme alone: the tile stayed marked and a save kept Glossy.
		await expect(undo(page)).toBeDisabled();
		expect(await appliedEffects(page)).toStrictEqual([]);
		expect(await savedEffects(page)).toBe('Office Theme');
		await redo.click();
		expect(await appliedEffects(page)).toStrictEqual(['glossy']);
		expect(await savedEffects(page)).toBe('Glossy');
	});

	test('hovering a scheme records no undo step', async ({ page }) => {
		await open(page);
		const root = await openGallery(page, 'effects');
		await root.locator('[data-gallery-item="riblet"]').hover();
		await page.waitForTimeout(400);
		await page.mouse.move(700, 600);
		await page.waitForTimeout(300);
		await expect(undo(page)).toBeDisabled();
	});
});

test.describe('Design > Variants', () => {
	test('says the four variants are computed from the theme colors', async ({ page }) => {
		await open(page);
		const root = page.locator('[data-ribbon-control="design.variants.gallery"]');
		await root.locator('[data-ribbon-gallery]').first().click();
		await expect(root.locator('.heading')).toHaveText('Computed from the theme colors');
	});
});
