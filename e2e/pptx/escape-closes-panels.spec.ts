/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Escape closes Outline view and the docked Format Background panel.
 *
 * Outline view covers the editor, and its close button was the only way back
 * to Normal view: no binding closed it on Escape. The rows had key handlers,
 * but opening the view leaves focus on the ribbon button behind the overlay.
 *
 * Design > Format Background opens a docked panel in VanillaJS and Svelte (the
 * others open the format pane instead, see `support/format-background.ts`). The VanillaJS panel had no Escape
 * handling at all, rendered without any styles, and its close button announced
 * itself as "Close find and replace" (copied from the find bar); the Svelte
 * panel handled Escape only with focus inside it, which opening never put
 * there.
 *
 * Run: bunx playwright test escape-closes-panels
 */
import { expect, test } from '@playwright/test';

import { loadDeck, openRibbonTab, SAMPLE_DECK } from './support/deck';
import { dockedFormatBackgroundGap } from './support/format-background';

test.use({ viewport: { width: 1440, height: 900 } });

test.describe('Escape closes panels', () => {
	test('Escape leaves Outline view for Normal view', async ({ page }) => {
		await loadDeck(page, SAMPLE_DECK);
		await openRibbonTab(page, 'View');
		await page.getByRole('button', { name: 'Outline View', exact: true }).first().click();
		const outline = page.locator('[data-pptx-outline-view]');
		await expect(outline).toBeVisible();

		await page.keyboard.press('Escape');
		await expect(outline, 'Escape must close Outline view, not only its close button').toBeHidden();
	});

	test('a docked Format Background panel is styled, named and closes on Escape', async ({
		page,
	}, testInfo) => {
		const gap = dockedFormatBackgroundGap(testInfo.project.name);
		test.skip(gap !== null, gap ?? '');

		await loadDeck(page, SAMPLE_DECK);
		await openRibbonTab(page, 'Design');
		const opener = page.getByRole('button', { name: 'Format Background', exact: true }).first();
		await opener.click();
		const panel = page.getByRole('dialog', { name: 'Format Background' });
		// Opened for real first, or every check below passes against a missing dialog.
		await expect(panel).toBeVisible();

		await expect(panel.getByRole('button', { name: 'Close find and replace' })).toHaveCount(0);
		const background = await panel.evaluate((node) => getComputedStyle(node).backgroundColor);
		expect(background, 'the panel must have its own background, not render unstyled').not.toMatch(
			/^(?:transparent|rgba\(0, 0, 0, 0\))$/u,
		);
		// Focus moves in on open, so Escape reaches the panel's own handler.
		await expect
			.poll(() => panel.evaluate((node) => node.contains(document.activeElement)))
			.toBe(true);

		await page.keyboard.press('Escape');
		await expect(panel).toBeHidden();
		await expect(opener, 'closing hands focus back to the ribbon command').toBeFocused();
	});
});
