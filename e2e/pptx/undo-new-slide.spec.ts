/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Undoing New Slide hands the stage back to a slide that exists.
 *
 * New Slide inserts after the slide on screen and shows the new one. When that
 * was the last slide, the new slide is the last too, so undoing it removes the
 * very slide the stage is showing. Two bindings restored the deck and kept the
 * old slide index, one past the new end, and the stage went blank (no elements,
 * no current thumbnail) until the user clicked a thumbnail.
 *
 * Run: bunx playwright test undo-new-slide
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { loadDeck, openRibbonTab, SAMPLE_DECK, slideElements, thumbnail } from './support/deck';

test.use({ viewport: { width: 1440, height: 900 } });

const slideCount = (page: Page): Promise<number> =>
	page.locator('[aria-label^="Go to slide"]').count();

test.describe('undo of New Slide', () => {
	test('undoing a slide added after the last slide shows the last slide again', async ({
		page,
	}) => {
		await loadDeck(page, SAMPLE_DECK);
		const before = await slideCount(page);
		await thumbnail(page, before).click();
		await expect(slideElements(page).first()).toBeVisible();
		const lastSlideElements = await slideElements(page).count();
		expect(lastSlideElements, 'the fixture deck has content on its last slide').toBeGreaterThan(0);

		await openRibbonTab(page, 'Home');
		await page
			.getByRole('button', { name: /^new slide$/iu })
			.first()
			.click();
		await expect.poll(() => slideCount(page)).toBe(before + 1);

		await page.getByRole('button', { name: /^undo/iu }).first().click();
		await expect.poll(() => slideCount(page)).toBe(before);
		await expect(
			slideElements(page),
			'the stage must show the (restored) last slide, not an index past the end of the deck',
		).toHaveCount(lastSlideElements);
	});
});
