/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
/**
 * A touch long-press opens the same context menus a right-click does, in every binding.
 *
 * iOS Safari fires no `contextmenu` for a long-press, so each binding installs the shared
 * touch-press recogniser on its viewer root and marks its menu hosts with
 * `data-pptx-touch-press`. The press is sent as touch pointer events from the page, which
 * every engine (WebKit included) delivers the same way, rather than as CDP touches.
 *
 * Run: bunx playwright test touch-long-press
 */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { stageElements } from './support/context-menu';
import { loadDeck, slideStage } from './support/deck';

test.use({ viewport: { width: 1440, height: 1300 } });

const menu = (page: Page): Locator => page.locator('pptx-ui-context-menu');

/** Hold a touch on the middle of `target` past the long-press delay, then lift it. */
async function longPress(target: Locator): Promise<void> {
	await target.evaluate(async (element) => {
		const box = element.getBoundingClientRect();
		const init: PointerEventInit = {
			bubbles: true,
			cancelable: true,
			composed: true,
			pointerId: 41,
			pointerType: 'touch',
			isPrimary: true,
			clientX: box.left + box.width / 2,
			clientY: box.top + box.height / 2,
		};
		element.dispatchEvent(new PointerEvent('pointerdown', init));
		await new Promise((resolve) => setTimeout(resolve, 700));
		element.dispatchEvent(new PointerEvent('pointerup', init));
	});
}

async function openDeck(page: Page): Promise<void> {
	await loadDeck(page);
	await slideStage(page).waitFor();
	await page.waitForTimeout(500);
}

test('a long-press on a slide thumbnail opens the slide menu', async ({ page }) => {
	await openDeck(page);
	await expect(menu(page)).toHaveCount(0);
	await longPress(page.locator('[aria-label="Go to slide 2"]').first());
	await expect(menu(page)).toHaveAttribute('data-pptx-slide-pane-context-menu', 'true');
	await expect(menu(page).getByRole('menuitem', { name: /^Duplicate/u })).toBeVisible();
});

test('a long-press on a shape opens the element menu', async ({ page }) => {
	await openDeck(page);
	await expect(menu(page)).toHaveCount(0);
	await longPress(stageElements(page).filter({ hasText: 'Product Overview' }).first());
	await expect(menu(page).getByRole('menu', { name: 'Context menu' })).toBeVisible();
	await expect(menu(page).getByRole('menuitem', { name: 'Copy', exact: true })).toBeVisible();
});

test('a short touch opens no menu', async ({ page }) => {
	await openDeck(page);
	const thumb = page.locator('[aria-label="Go to slide 2"]').first();
	await thumb.evaluate((element) => {
		const box = element.getBoundingClientRect();
		const init: PointerEventInit = {
			bubbles: true,
			composed: true,
			pointerId: 42,
			pointerType: 'touch',
			isPrimary: true,
			clientX: box.left + box.width / 2,
			clientY: box.top + box.height / 2,
		};
		element.dispatchEvent(new PointerEvent('pointerdown', init));
		element.dispatchEvent(new PointerEvent('pointerup', init));
	});
	await page.waitForTimeout(800);
	await expect(menu(page)).toHaveCount(0);
});
