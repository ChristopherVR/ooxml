/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Picture Format > Size > Crop, end to end in every binding.
 *
 * Selecting a picture switches the ribbon to Picture Format, so crop has to be
 * reachable there and not only from Home > Arrange. The group is shared
 * (`CONTEXTUAL_TAB_GROUPS.pictureFormat`, the `picture-size` Home strip); this
 * spec proves each demo renders it, wires it to the same crop controller and
 * honours the customisation ids.
 *
 * Fixture: `merge-crop.pptx`, whose picture is "Crop Pic".
 *
 * Run: PPTX_E2E_PORT_OFFSET=100 bunx playwright test picture-format-crop --project=react
 */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { elementsOfType, elementWithText, fixture, loadDeck, selectElement } from './support/deck';

test.use({ viewport: { width: 1440, height: 900 } });

const DECK = fixture('merge-crop.pptx');

function sizeGroup(page: Page): Locator {
	return page.locator('[data-ribbon-group="pictureFormat.size"]').first();
}

function sizeCrop(page: Page, hook: 'crop' | 'crop-menu'): Locator {
	return sizeGroup(page).locator(`[data-pptx-ribbon-control="${hook}"]`).first();
}

function cropOverlay(page: Page): Locator {
	return page.locator('[data-pptx-crop-overlay="true"]');
}

async function openPictureFormat(page: Page, path = '/'): Promise<void> {
	await loadDeck(page, DECK, path);
	await elementWithText(page, 'Merge B').waitFor();
	await selectElement(page, elementsOfType(page, 'image').first());
	await page.locator('[data-ribbon-contextual-tab="pictureFormat"]').first().click();
}

test('picture Format > Size > Crop starts and ends crop mode', async ({ page }) => {
	await openPictureFormat(page);
	await expect(sizeGroup(page)).toBeVisible();
	await expect(sizeCrop(page, 'crop')).toHaveAttribute('aria-pressed', 'false');

	await sizeCrop(page, 'crop').click();
	await expect(cropOverlay(page)).toBeVisible();
	await expect(sizeCrop(page, 'crop')).toHaveAttribute('aria-pressed', 'true');

	await sizeCrop(page, 'crop').click();
	await expect(cropOverlay(page)).toHaveCount(0);
});

test('picture Format > Size > Crop to Aspect Ratio 1:1 squares the picture', async ({ page }) => {
	await openPictureFormat(page);
	await sizeCrop(page, 'crop-menu').click();
	await page.locator('[data-pptx-crop-aspect="1:1"]:visible').first().click();
	const picture = elementsOfType(page, 'image').first();
	await expect
		.poll(async () => {
			const box = await picture.boundingBox();
			return box ? Math.abs(box.width - box.height) : Number.POSITIVE_INFINITY;
		})
		.toBeLessThan(3);
});

test('a host can hide the Size group, or crop everywhere', async ({ page }) => {
	const query = (customization: object) =>
		`/?customization=${encodeURIComponent(JSON.stringify(customization))}`;
	await openPictureFormat(page, query({ ribbon: { hiddenGroups: ['pictureFormat.size'] } }));
	await expect(page.locator('[data-ribbon-group="pictureFormat.pictureStyles"]')).toBeVisible();
	await expect(sizeGroup(page)).toBeHidden();

	await openPictureFormat(page, query({ ribbon: { hiddenButtons: ['crop'] } }));
	await expect(sizeGroup(page)).toBeVisible();
	await expect(sizeCrop(page, 'crop')).toBeHidden();
	await expect(sizeCrop(page, 'crop-menu')).toBeHidden();
});
