/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The host `limits` option, driven through every binding by the demos'
 * `?maxSlides=` query option (see `demos/pptx/shared/demo-root-options.ts`).
 * A deck over the limit fails with the localized reason; a deck within it
 * renders as usual. The viewers show the reason in their load-error alert;
 * the demos that report a failed load through the viewer's error callback
 * show that same message on their upload screen instead, so the spec looks
 * for the text rather than one container.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { elementWithText, resetTabSession, slideStage, thumbnail } from './support/deck';
import { LOAD_LIMITS_DECK_SLIDES, loadLimitsDeck } from './support/load-limits-deck';

const OVER_LIMIT_MESSAGE = `This presentation has ${LOAD_LIMITS_DECK_SLIDES} slides, more than the ${
	LOAD_LIMITS_DECK_SLIDES - 1
} this viewer accepts.`;

async function openDeck(page: Page, query: string): Promise<void> {
	await resetTabSession(page);
	await page.goto(`/${query}`);
	await page.locator('#file-input').setInputFiles({
		name: 'load-limits.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: await loadLimitsDeck(),
	});
}

test.describe('limits host option', () => {
	test('a deck over maxSlides fails with the localized limit and slide count', async ({ page }) => {
		await openDeck(page, `?maxSlides=${LOAD_LIMITS_DECK_SLIDES - 1}`);
		await expect(
			page.getByText(OVER_LIMIT_MESSAGE, { exact: true }).filter({ visible: true }),
		).toHaveCount(1);
		await expect(elementWithText(page, 'Load limits slide 1')).toHaveCount(0);
	});

	test('a deck within maxSlides renders every slide', async ({ page }) => {
		await openDeck(page, `?maxSlides=${LOAD_LIMITS_DECK_SLIDES}`);
		await expect(slideStage(page)).toBeVisible();
		await expect(elementWithText(page, 'Load limits slide 1').first()).toBeVisible();
		await expect(thumbnail(page, LOAD_LIMITS_DECK_SLIDES)).toBeVisible();
		await expect(page.getByText('this viewer accepts')).toHaveCount(0);
	});
});
