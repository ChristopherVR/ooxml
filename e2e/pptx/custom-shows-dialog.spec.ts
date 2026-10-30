/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The Custom Shows dialog (Slide Show > Custom Slide Show): Copy, Remove and Show.
 *
 * The dialog's logic is one shared controller, so this drives only what every binding must
 * show: the saved-shows list, Copy adding "Copy of ...", and Show starting the selected show so
 * that only its slides play. `header-footer-shows.pptx` defines "Short Show" (slides 1 and 3)
 * and "Reverse" (3, 2, 1).
 *
 * Run: bunx playwright test custom-shows-dialog
 */
import { expect, test } from '@playwright/test';

import { fixture, resetTabSession } from './support/deck';

test.use({ viewport: { width: 1440, height: 900 } });

const DECK = fixture('header-footer-shows.pptx');

test.describe('custom shows dialog', () => {
	test('lists the deck shows, copies one and starts another', async ({ page }) => {
		await resetTabSession(page);
		await page.goto('/');
		await page.locator('#file-input').setInputFiles(DECK);
		await page.locator('[data-element-id]').filter({ hasText: 'Alpha Slide' }).first().waitFor();

		await page.getByRole('tab', { name: 'Slide Show' }).first().click();
		await page.getByRole('button', { name: 'Custom Slide Show' }).first().click();

		const saved = page.getByRole('listbox', { name: 'Saved Shows' });
		await expect(saved.getByRole('option', { name: 'Short Show' })).toBeVisible();
		await expect(saved.getByRole('option', { name: 'Reverse' })).toBeVisible();

		await saved.getByRole('option', { name: 'Short Show' }).click();
		await page.getByRole('button', { name: 'Copy', exact: true }).last().click();
		await expect(saved.getByRole('option', { name: 'Copy of Short Show' })).toBeVisible();

		await saved.getByRole('option', { name: 'Reverse' }).click();
		await page.getByRole('button', { name: 'Show', exact: true }).last().click();

		// The dialog committed and closed, and "Reverse" opens on slide 3.
		await expect(saved).toBeHidden();
		await expect(
			page
				.locator('[data-element-id]')
				.filter({ hasText: 'Gamma Slide' })
				.and(page.locator(':visible'))
				.first(),
		).toBeVisible({ timeout: 10_000 });
	});
});
