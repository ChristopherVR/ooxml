/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * One Undo takes back one Home > Layout pick, and one New Slide from a layout.
 *
 * `PptxHandler.applyLayoutToSlide` replaces the slide in the array it is given. Vue and Svelte
 * handed it their live slides, so the undo step they recorded after the call already held the new
 * layout: the first Undo after a layout pick changed nothing, and every later one landed a step
 * late. Svelte and vanilla also recorded a second step when a new slide's layout arrived, so one
 * Undo after New Slide from a layout left the new slide in the deck.
 *
 * Run: bunx playwright test layout-undo
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { loadDeck, openRibbonTab, SAMPLE_DECK } from './support/deck';

test.use({ viewport: { width: 1440, height: 900 } });

const MENU_TIMEOUT_MS = 15_000;

const slideCount = (page: Page): Promise<number> =>
	page.locator('[aria-label^="Go to slide"]').count();

const undoButton = (page: Page) => page.getByRole('button', { name: /^undo/iu }).first();

function layoutEntries(page: Page) {
	return page.locator('[data-testid="layout-gallery-menu"]:visible button');
}

async function openLayoutMenu(page: Page): Promise<void> {
	await openRibbonTab(page, 'Home');
	const trigger = page.getByRole('button', { name: /^layout$/iu }).first();
	await trigger.hover();
	await trigger.click().catch(() => {
		// Angular's menu opens on hover; the click can land on the panel now covering it.
	});
	await expect(layoutEntries(page).first()).toBeVisible({ timeout: MENU_TIMEOUT_MS });
}

async function closeMenu(page: Page): Promise<void> {
	await page.keyboard.press('Escape');
	await page.mouse.move(0, 0);
}

/** The layout the Layout menu marks as the slide's own. */
async function markedLayout(page: Page): Promise<string | undefined> {
	await openLayoutMenu(page);
	const marked = await layoutEntries(page)
		.and(page.locator('[aria-current="true"]'))
		.allInnerTexts();
	await closeMenu(page);
	return marked[0]?.trim().split('\n')[0];
}

async function applyLayout(page: Page, name: string): Promise<void> {
	await openLayoutMenu(page);
	await layoutEntries(page).filter({ hasText: name }).first().click();
	await closeMenu(page);
	await expect.poll(() => markedLayout(page), { timeout: MENU_TIMEOUT_MS }).toBe(name);
}

test.describe('undo of a layout change', () => {
	test('one Undo takes back one Home > Layout pick', async ({ page }) => {
		await loadDeck(page, SAMPLE_DECK);
		await expect.poll(() => markedLayout(page)).toBe('Title Slide');

		await applyLayout(page, 'Title and Content');
		await applyLayout(page, 'Two Content');

		await undoButton(page).click();
		await expect
			.poll(() => markedLayout(page), { message: 'the first Undo takes back Two Content' })
			.toBe('Title and Content');
		await undoButton(page).click();
		await expect
			.poll(() => markedLayout(page), { message: 'the second Undo takes back Title and Content' })
			.toBe('Title Slide');
	});

	test('one Undo takes back New Slide from a layout', async ({ page }) => {
		await loadDeck(page, SAMPLE_DECK);
		const before = await slideCount(page);

		await openRibbonTab(page, 'Home');
		await page
			.locator('[data-ribbon-control="home.slides.newSlide"] [data-pptx-chrome="split-caret"]')
			.click();
		// New Slide's menu lists the layouts in the same gallery as Home > Layout.
		await layoutEntries(page).filter({ hasText: 'Title and Content' }).first().click();
		await expect.poll(() => slideCount(page)).toBe(before + 1);
		// The core re-maps the new slide asynchronously, and no binding shows when it lands (an
		// empty slide on this layout draws nothing in edit view). Undo only after it has, or a
		// second step recorded on its arrival would go unseen.
		await page.waitForTimeout(2_000);

		await undoButton(page).click();
		await expect
			.poll(() => slideCount(page), { message: 'one Undo removes the new slide' })
			.toBe(before);
		await expect(undoButton(page)).toBeDisabled();
	});
});
