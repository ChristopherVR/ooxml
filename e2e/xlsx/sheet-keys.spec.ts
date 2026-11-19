import { expect, test, type Page } from '@playwright/test';
import { grid, nameBoxValue, openSample, SAMPLE_SHEETS, sheetTabs } from './helpers';

const selectedSheet = (page: Parameters<typeof sheetTabs>[0]) =>
	sheetTabs(page).locator('[role="tab"][aria-selected="true"]');

/** What a screen reader reads for the focused cell: the grid's focus sink label. */
const cellLabel = (page: Page) => grid(page).locator('textarea').first().getAttribute('aria-label');

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

test('the cell label reads the sheet Ctrl+PageDown switched to', async ({ page }) => {
	await openSample(page);
	await grid(page).click({ position: { x: 120, y: 60 } });
	const first = await cellLabel(page);
	await page.keyboard.press('Control+PageDown');
	await expect(selectedSheet(page)).toHaveText(SAMPLE_SHEETS[1]);
	// Moving away and back describes the same cell on the new sheet from scratch.
	await page.keyboard.press('ArrowDown');
	await page.keyboard.press('ArrowUp');
	const fresh = await cellLabel(page);
	expect(fresh).not.toBe(first);
	await page.keyboard.press('Control+PageUp');
	await page.keyboard.press('Control+PageDown');
	await expect.poll(() => cellLabel(page)).toBe(fresh);
});

/** The shell region (its shadow part) that holds keyboard focus. */
const focusedRegion = (page: Page) =>
	page.evaluate(() => {
		const active = document.querySelector('xlsx-editor')?.shadowRoot?.activeElement;
		const region = active?.closest(
			['ribbon', 'formula-bar', 'grid', 'sheet-tabs', 'status-bar']
				.map((name) => `[part~="${name}"]`)
				.join(','),
		);
		return region?.getAttribute('part') ?? 'none';
	});

for (const key of ['F6', 'Shift+F6']) {
	test(`${key} visits every region and comes back to the grid`, async ({ page }) => {
		await openSample(page);
		await grid(page).click({ position: { x: 120, y: 60 } });
		expect(await focusedRegion(page)).toContain('grid');
		const seen: string[] = [];
		for (let i = 0; i < 5; i++) {
			await page.keyboard.press(key);
			seen.push(await focusedRegion(page));
		}
		// Five regions in a cycle: each press lands on a new one, the status bar included, and the
		// fifth press is back on the grid.
		expect(new Set(seen).size, seen.join(' > ')).toBe(5);
		expect(seen.some((part) => part.includes('status-bar'))).toBe(true);
		expect(seen.at(-1)).toContain('grid');
	});
}
