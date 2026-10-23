/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
import { test, expect } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { elementsOfType, fixture, inspector, loadDeck, selectElement } from './support/deck';

/**
 * The chart "Edit Data" sheet. Every binding mounts the one shared
 * `pptx-ui-chart-data-sheet` element and only commits what it emits, so this
 * spec drives the same sheet in each project: a typed value commits on Enter
 * and moves down a row, Escape restores a cell, and a pasted tab-separated
 * block grows the chart by a category the canvas then draws.
 */
const CHART_DECK = fixture('chart-title-runs.pptx');

async function openSheet(page: Page): Promise<Locator> {
	await page.setViewportSize({ width: 1280, height: 900 });
	await loadDeck(page, CHART_DECK);
	const chart = elementsOfType(page, 'chart').first();
	await chart.waitFor();
	if (
		!(await inspector(page)
			.isVisible()
			.catch(() => false))
	) {
		const toggle = page.getByRole('button', { name: 'Toggle inspector panel' });
		if (await toggle.isVisible().catch(() => false)) {
			await toggle.click();
		}
	}
	await selectElement(page, chart);
	const sheet = inspector(page).locator('pptx-ui-chart-data-sheet');
	await sheet.waitFor({ state: 'attached', timeout: 15_000 });
	await sheet.scrollIntoViewIfNeeded();
	return sheet;
}

const cell = (sheet: Locator, row: number, col: number): Locator =>
	sheet.locator(`input[data-row="${row}"][data-col="${col}"]`);

test.describe('chart data sheet', () => {
	test('commits typed values, moves with the keyboard and pastes blocks', async ({ page }) => {
		test.setTimeout(120_000);
		const sheet = await openSheet(page);

		// Type into the first value cell and commit with Enter: the value sticks
		// after the binding re-renders and focus moves one row down.
		const first = cell(sheet, 1, 1);
		await first.click();
		await first.fill('123');
		await first.press('Enter');
		await expect(cell(sheet, 1, 1)).toHaveValue('123');
		await expect(cell(sheet, 2, 1)).toBeFocused();

		// Escape restores the cell's text instead of committing it.
		const second = cell(sheet, 2, 1);
		const before = await second.inputValue();
		await second.fill('999');
		await second.press('Escape');
		await expect(second).toHaveValue(before);

		// Text that is not a number is refused and the cell restored.
		await second.fill('abc');
		await second.press('Enter');
		await expect(cell(sheet, 2, 1)).toHaveValue(before);

		// Paste a block over the last category row: its second row adds a
		// category, whose label the chart then draws.
		const rows = await sheet.locator('tbody tr').count();
		const last = cell(sheet, rows, 0);
		await last.evaluate((input) => {
			const data = new DataTransfer();
			data.setData('text/plain', 'Pasted A\t5\r\nPasted B\t6\r\n');
			input.dispatchEvent(
				new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
			);
		});
		await expect(sheet.locator('tbody tr')).toHaveCount(rows + 1);
		await expect(cell(sheet, rows + 1, 0)).toHaveValue('Pasted B');
		await expect(cell(sheet, rows + 1, 1)).toHaveValue('6');
		const chartSvg = elementsOfType(page, 'chart').first().locator('svg').first();
		await expect
			.poll(async () => chartSvg.evaluate((svg) => svg.textContent ?? ''), { timeout: 10_000 })
			.toContain('Pasted B');
	});
});
