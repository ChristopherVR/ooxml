/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Review > Comments: Previous, Next and Delete across every framework demo.
 *
 * The runner (`runReviewCommentCommand`) is shared and unit tested; this pins the wiring each
 * binding does around it: its slides, its slide navigator, its comments pane (rows tagged
 * `data-comment-id`, the current one `data-comment-active`) and its history-aware writer.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import {
	addCommentViaContextMenu,
	commentTextVisible,
	COMPOSE_BOX_SELECTOR,
} from './support/comments';
import { loadDeck, openRibbonTab, ribbon, SAMPLE_DECK, slideElements } from './support/deck';

test.use({ viewport: { width: 1440, height: 900 } });

const FIRST = 'Navigate to me first';
const SECOND = 'Navigate to me second';

const button = (page: Page, name: string) =>
	ribbon(page).getByRole('button', { name, exact: true }).first();

test('Previous and Next walk the comments and Delete removes the current one', async ({ page }) => {
	await loadDeck(page, SAMPLE_DECK);
	const shape = slideElements(page).filter({ hasText: 'Product Overview' }).first();
	await shape.waitFor();
	await page.waitForTimeout(400);
	await addCommentViaContextMenu(page, shape, FIRST);
	// The pane that took the first comment is still open in most bindings: compose in it.
	const compose = page.locator(COMPOSE_BOX_SELECTOR).first();
	if (await compose.isVisible()) {
		await compose.fill(SECOND);
		await page.locator('button:visible:text-is("Add Comment")').last().click();
		await page.waitForTimeout(500);
	} else {
		await addCommentViaContextMenu(page, shape, SECOND);
	}

	await openRibbonTab(page, 'Review');
	// Nothing is current yet, so Delete has nothing to act on.
	await expect(button(page, 'Delete')).toBeDisabled();
	await expect(button(page, 'Next')).toBeEnabled();

	await button(page, 'Next').click();
	await expect(page.locator('[data-comment-active]').first()).toContainText(FIRST);
	await button(page, 'Next').click();
	await expect(page.locator('[data-comment-active]').first()).toContainText(SECOND);
	// Next wraps to the first comment, Previous goes back.
	await button(page, 'Next').click();
	await expect(page.locator('[data-comment-active]').first()).toContainText(FIRST);
	await button(page, 'Previous').click();
	await expect(page.locator('[data-comment-active]').first()).toContainText(SECOND);

	await expect(button(page, 'Delete')).toBeEnabled();
	await button(page, 'Delete').click();
	await expect.poll(() => commentTextVisible(page, SECOND)).toBe(false);
	await expect.poll(() => commentTextVisible(page, FIRST)).toBe(true);
});

test('Delete All Comments removes every comment in the presentation', async ({ page }) => {
	await loadDeck(page, SAMPLE_DECK);
	const shape = slideElements(page).filter({ hasText: 'Product Overview' }).first();
	await shape.waitFor();
	await page.waitForTimeout(400);
	await addCommentViaContextMenu(page, shape, FIRST);

	await openRibbonTab(page, 'Review');
	await button(page, 'Delete All Comments').click();
	await expect.poll(() => commentTextVisible(page, FIRST)).toBe(false);
	await expect(button(page, 'Next')).toBeDisabled();
});
