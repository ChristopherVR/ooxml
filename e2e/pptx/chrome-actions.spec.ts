/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
/**
 * The ribbon tab row's right end, framework-neutral: every binding draws Comments and Share
 * with the shared `pptx-ui-ribbon-actions` (the `office-ui-ribbon-actions` Word and Excel put
 * at the same place), right of the tabs, after Record, in that order. PowerPoint 365 has no
 * editing-mode selector there, so none is drawn, and Comments has left the quick-access row.
 */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { COMPOSE_BOX_SELECTOR } from './support/comments';
import { loadDeck, slideStage } from './support/deck';

test.use({ viewport: { width: 1440, height: 900 } });

const tabRow = (page: Page): Locator => page.locator('[data-pptx-chrome="ribbon-tabs"]').first();
const actions = (page: Page): Locator => tabRow(page).locator('pptx-ui-ribbon-actions');
const comments = (page: Page): Locator =>
	actions(page).getByRole('button', { name: 'Comments', exact: true });
const share = (page: Page): Locator =>
	actions(page).getByRole('button', { name: 'Share', exact: true });

async function open(page: Page): Promise<void> {
	await loadDeck(page);
	await slideStage(page).waitFor();
	await expect(actions(page)).toHaveCount(1);
}

test('Comments and Share sit right of the tabs, after Record, in Office order', async ({
	page,
}, info) => {
	await open(page);
	await expect(comments(page)).toBeVisible();
	await expect(share(page)).toBeVisible();
	await expect(comments(page)).toHaveText('Comments');
	await expect(share(page)).toHaveText('Share');
	// No editing-mode selector on PowerPoint's tab row.
	await expect(actions(page).getByRole('combobox')).toHaveCount(0);
	// Comments moved off the quick-access row; Share is not duplicated there.
	const primary = page.locator('[data-pptx-chrome="ribbon-primary"]').first();
	await expect(primary.getByRole('button', { name: 'Comments', exact: true })).toHaveCount(0);
	await expect(primary.getByRole('button', { name: 'Share', exact: true })).toHaveCount(0);

	const tabs = tabRow(page).getByRole('tab');
	const lastTab = (await tabs.last().boundingBox())!;
	const [commentsBox, shareBox] = await Promise.all(
		[comments(page), share(page)].map(async (control) => (await control.boundingBox())!),
	);
	for (const box of [commentsBox!, shareBox!]) {
		const middle = box.y + box.height / 2;
		expect(middle).toBeGreaterThan(lastTab.y);
		expect(middle).toBeLessThan(lastTab.y + lastTab.height);
		expect(box.x).toBeGreaterThan(lastTab.x + lastTab.width);
	}
	expect(commentsBox!.x).toBeLessThan(shareBox!.x);
	const record = tabRow(page).getByRole('button', { name: 'Record', exact: true });
	if (await record.count()) {
		expect((await record.boundingBox())!.x).toBeLessThan(commentsBox!.x);
	}
	await page.screenshot({
		path: info.outputPath('tab-row.png'),
		clip: { x: 0, y: 0, width: 1440, height: 140 },
	});
});

test('Share opens the Share dialog and Comments opens the comments UI', async ({ page }) => {
	await open(page);
	// Share first: Vanilla's comments pane overlays the whole viewer, tab row included.
	await share(page).click();
	const dialog = page
		.getByRole('dialog')
		.filter({ has: page.locator('input') })
		.last();
	await expect(dialog).toBeVisible();
	await page.keyboard.press('Escape');
	await expect(dialog).toBeHidden();
	await comments(page).click();
	await expect(page.locator(COMPOSE_BOX_SELECTOR).first()).toBeVisible();
});
