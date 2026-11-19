/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The thumbnail rail of a deck above the virtualization threshold, in every
 * binding. Paging with the keyboard from a thumbnail must keep focus on the
 * current slide's thumbnail and keep it in view: the rail drops rows that
 * scroll out, which used to drop focus to the page, after which the keys
 * scrolled the rail instead of changing slides. Thumbnails that scroll into
 * view are drawn once scrolling pauses, so after scrolling the rail shows the
 * slides it scrolled to.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { resetTabSession, slideStage, thumbnail } from './support/deck';
import { largeDeck } from './support/large-deck';

const PAGES = 25;

async function openLargeDeck(page: Page): Promise<void> {
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles({
		name: 'large-deck.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: await largeDeck(),
	});
	await expect(slideStage(page)).toBeVisible();
	await expect(thumbnail(page, 1)).toBeVisible();
}

test.describe('large deck thumbnail rail', () => {
	test('paging from a thumbnail keeps focus on the current slide', async ({ page }) => {
		await openLargeDeck(page);
		await thumbnail(page, 1).click();
		for (let step = 0; step < PAGES; step += 1) {
			await page.keyboard.press('PageDown');
		}
		const current = thumbnail(page, PAGES + 1);
		await expect(current).toBeFocused();
		await expect(current).toHaveAttribute('aria-current', 'true');
		await expect(current).toBeInViewport();
	});

	test('rows scrolled into view show their slides once scrolling stops', async ({ page }) => {
		await openLargeDeck(page);
		const list = page.locator('[data-pptx-chrome="slide-list"]').first();
		// Ten scroll steps in a row, then a pause: the last position is what shows.
		const step = await thumbnail(page, 2).evaluate(
			(row, first) => row.getBoundingClientRect().top - first,
			await thumbnail(page, 1).evaluate((row) => row.getBoundingClientRect().top),
		);
		await list.evaluate(async (el, rowStep) => {
			for (let index = 1; index <= 10; index += 1) {
				el.scrollTop = index * 4 * rowStep;
				await new Promise((resolve) => {
					requestAnimationFrame(resolve);
				});
			}
		}, step);
		await expect(thumbnail(page, 1)).toHaveCount(0);
		await expect(thumbnail(page, 42)).toBeInViewport();
		await expect(thumbnail(page, 42)).toContainText('Large deck slide 42');
	});
});
