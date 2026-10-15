/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The Table Layout contextual tab, run identically against every framework demo.
 *
 * It appears next to Table Design when a table is selected and its commands act
 * on the selected cells: insert a row below, merge two cells and split them
 * again, align the cell text. The unit tests (`table-layout-commands.test.ts`)
 * pin the edits; this pins the wiring of the selection and the gallery path in
 * each binding.
 */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { selectTableCell } from './support/context-menu';
import { fixture, loadDeck, ribbon } from './support/deck';

test.use({ viewport: { width: 1440, height: 900 } });

const table = (page: Page): Locator =>
	page.locator('[aria-roledescription="slide"]').first().locator('table').first();

const cellAt = (page: Page, row: number, column: number): Locator =>
	table(page).locator('tr').nth(row).locator('td').nth(column);

const command = (page: Page, id: string): Locator =>
	ribbon(page).locator(`[data-ribbon-control="${id}"] button`).first();

async function openTableLayout(page: Page): Promise<void> {
	await loadDeck(page, fixture('table-styling.pptx'));
	await selectTableCell(page, cellAt(page, 0, 0));
	await page.locator('[data-ribbon-contextual-tab="tableLayout"]').first().click();
	await page.waitForTimeout(200);
}

test("Table Layout shows next to Table Design with PowerPoint's groups", async ({ page }) => {
	await loadDeck(page, fixture('table-styling.pptx'));
	await selectTableCell(page, cellAt(page, 0, 0));
	await expect(page.locator('[data-ribbon-contextual-tab="tableDesign"]').first()).toBeVisible();
	await page.locator('[data-ribbon-contextual-tab="tableLayout"]').first().click();
	const labels = await page
		.locator('pptx-ui-ribbon-group[data-ribbon-group^="tableLayout."]')
		.evaluateAll((groups) => groups.map((group) => group.getAttribute('label')));
	expect(labels).toEqual([
		'Table',
		'Rows & Columns',
		'Merge',
		'Cell Size',
		'Alignment',
		'Table Size',
		'Arrange',
	]);
});

test('Insert Below adds a row under the selected cell', async ({ page }) => {
	await openTableLayout(page);
	const rows = await table(page).locator('tr').count();
	await command(page, 'tableLayout.rowsColumns.insertBelow').click();
	await expect(table(page).locator('tr')).toHaveCount(rows + 1);
});

test('Merge Cells is disabled for one cell and merges a two-cell range', async ({ page }) => {
	await openTableLayout(page);
	await expect(command(page, 'tableLayout.merge.mergeCells')).toBeDisabled();
	const columns = await table(page).locator('tr').first().locator('td').count();
	await cellAt(page, 0, 1).click({ modifiers: ['Shift'] });
	await command(page, 'tableLayout.merge.mergeCells').click();
	await expect(table(page).locator('tr').first().locator('td')).toHaveCount(columns - 1);
});

test('Cell Margins and Select stay disabled with a reason', async ({ page }) => {
	await openTableLayout(page);
	for (const id of ['tableLayout.alignment.cellMargins', 'tableLayout.table.select']) {
		const button = command(page, id);
		await expect(button).toBeDisabled();
		await expect(button).toHaveAttribute('title', /.+/u);
	}
});
