/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { loadDeck, ribbonTab, slideStage } from './support/deck';

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
		// One undo step puts the slide back (the other slides were never touched; the
		// unit and core tests pin that).
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
