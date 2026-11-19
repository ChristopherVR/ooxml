import { expect, test } from '@playwright/test';
import { nameBoxValue, openSample, SAMPLE_SHEETS, sheetTabs } from './helpers';

const selectedSheet = (page: Parameters<typeof sheetTabs>[0]) =>
	sheetTabs(page).locator('[role="tab"][aria-selected="true"]');

test('Ctrl+PageDown and Ctrl+PageUp switch sheets and leave the cell alone', async ({ page }) => {
	await openSample(page);
	await expect(selectedSheet(page)).toHaveText(SAMPLE_SHEETS[0]);
	const cell = await nameBoxValue(page);

	await page.keyboard.press('Control+PageDown');
	await expect(selectedSheet(page)).toHaveText(SAMPLE_SHEETS[1]);
	expect(await nameBoxValue(page)).toBe(cell);

	await page.keyboard.press('Control+PageUp');
	await expect(selectedSheet(page)).toHaveText(SAMPLE_SHEETS[0]);
	expect(await nameBoxValue(page)).toBe(cell);

	// Without Ctrl the key still moves the cell by a screen.
	await page.keyboard.press('PageDown');
	await expect(selectedSheet(page)).toHaveText(SAMPLE_SHEETS[0]);
	await expect.poll(() => nameBoxValue(page)).not.toBe(cell);
});
