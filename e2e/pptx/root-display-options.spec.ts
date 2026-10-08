/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The root display options every binding accepts: `showToolbar`,
 * `showThumbnails` and `initialSlide`.
 *
 * The defaults and the clamping are decided once in
 * `src/ui/src/pptx/render/viewer-root-options.ts` and pinned there by unit
 * tests; what this spec catches is a binding whose template never consults
 * the resolved options, so the unit suites are green and the ribbon is still
 * on screen. The demos read `?showToolbar=0`, `?showThumbnails=0` and
 * `?initialSlide=<n>` (see `demos/pptx/shared/demo-root-options.ts`), so the
 * same steps run unchanged against all five demos.
 */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { loadDeck, ribbon, SAMPLE_DECK, slideStage, thumbnail } from './support/deck';

test.use({ viewport: { width: 1440, height: 900 } });

/** The sample deck's slide count. */
const SAMPLE_SLIDES = 7;

const titleBar = (page: Page): Locator => page.locator('pptx-ui-title-bar:visible');
const statusBar = (page: Page): Locator => page.locator('pptx-ui-status-bar:visible');

async function expectCounter(page: Page, slideNumber: number): Promise<void> {
	await expect(
		statusBar(page).getByText(new RegExp(`^Slide ${slideNumber} of ${SAMPLE_SLIDES}$`, 'u')),
	).toBeVisible();
}

test.describe('root display options', () => {
	test('the defaults show the chrome and the slide pane and open on slide 1', async ({ page }) => {
		await loadDeck(page, SAMPLE_DECK);
		await expect(titleBar(page)).toHaveCount(1);
		await expect(ribbon(page)).toBeVisible();
		await expect(statusBar(page)).toHaveCount(1);
		await expect(thumbnail(page, 1)).toBeVisible();
		await expectCounter(page, 1);
	});

	test('showToolbar=false hides the title bar, ribbon and status bar', async ({ page }) => {
		await loadDeck(page, SAMPLE_DECK, '/?showToolbar=0');
		await expect(slideStage(page)).toBeVisible();
		await expect(titleBar(page)).toHaveCount(0);
		await expect(ribbon(page)).toHaveCount(0);
		await expect(statusBar(page)).toHaveCount(0);
		// The slide pane is a separate option and stays.
		await expect(thumbnail(page, 1)).toBeVisible();
	});

	test('showThumbnails=false hides the slide pane and keeps the chrome', async ({ page }) => {
		await loadDeck(page, SAMPLE_DECK, '/?showThumbnails=0');
		await expect(slideStage(page)).toBeVisible();
		await expect(thumbnail(page, 1)).toHaveCount(0);
		await expect(titleBar(page)).toHaveCount(1);
		await expect(ribbon(page)).toBeVisible();
		await expectCounter(page, 1);
	});

	test('initialSlide opens the deck on that zero-based slide', async ({ page }) => {
		await loadDeck(page, SAMPLE_DECK, '/?initialSlide=3');
		await expectCounter(page, 4);
	});

	test('initialSlide is clamped into the deck', async ({ page }) => {
		await loadDeck(page, SAMPLE_DECK, '/?initialSlide=99');
		await expectCounter(page, SAMPLE_SLIDES);
		await loadDeck(page, SAMPLE_DECK, '/?initialSlide=-4');
		await expectCounter(page, 1);
	});
});
