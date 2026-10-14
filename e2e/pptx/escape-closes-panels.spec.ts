/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Escape closes Outline view and the docked Format Background panel.
 *
 * Outline view covers the editor, and its close button was the only way back
 * to Normal view: no binding closed it on Escape. The rows had key handlers,
 * but opening the view leaves focus on the ribbon button behind the overlay.
 *
 * Design > Format Background opens a docked panel in the bindings that have one
 * (the others open the format pane instead). The VanillaJS panel had no Escape
 * handling at all, rendered without any styles, and its close button announced
 * itself as "Close find and replace" (copied from the find bar); the Svelte
 * panel handled Escape only with focus inside it, which opening never put
 * there.
 *
 * Run: bunx playwright test escape-closes-panels
 */
import { expect, test } from '@playwright/test';

import { loadDeck, openRibbonTab, SAMPLE_DECK } from './support/deck';

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
	}) => {
		await loadDeck(page, SAMPLE_DECK);
		await openRibbonTab(page, 'Design');
		await page.getByRole('button', { name: 'Format Background', exact: true }).first().click();
		await page.waitForTimeout(300);
		// Bindings without a docked panel open the format pane: no dialog, nothing to check.
		const panel = page.getByRole('dialog', { name: 'Format Background' });

		await expect(panel.getByRole('button', { name: 'Close find and replace' })).toHaveCount(0);
		const unstyled = await panel.evaluateAll(
			(nodes) =>
				nodes
					.map((node) => getComputedStyle(node))
					.filter(
						(style) =>
							style.backgroundColor === 'rgba(0, 0, 0, 0)' ||
							style.backgroundColor === 'transparent',
					).length,
		);
		expect(unstyled, 'the panel must have its own background, not render unstyled').toBe(0);

		await page.keyboard.press('Escape');
		await expect(panel).toBeHidden();
	});
});
