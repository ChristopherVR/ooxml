import { expect, test } from '@playwright/test';
import { editor, FRAMEWORKS, goToCell, grid, newWorkbook, typeInActiveCell } from './helpers';

/** What a screen reader explores: the grid's rows and cells, and the cell the focus points at. */
for (const framework of FRAMEWORKS)
	test(`the grid is a table of rows and cells in ${framework}`, async ({ page }) => {
		await newWorkbook(page, framework);
		await goToCell(page, 'B2');
		await typeInActiveCell(page, '100');
		const table = editor(page).getByRole('grid', { name: 'Spreadsheet grid' });
		const cell = table.getByRole('gridcell', { name: '100', exact: true });
		// B2 is row 2, column 2, as Excel reports it: the headers take no index.
		await expect(cell).toHaveAttribute('aria-colindex', '2');
		await expect(cell.locator('..')).toHaveAttribute('aria-rowindex', '2');
		await expect(table.getByRole('columnheader', { name: 'B', exact: true })).toHaveAttribute(
			'aria-colindex',
			'2',
		);

		await page.keyboard.press('ArrowUp');
		const sink = grid(page).locator('textarea').first();
		await expect(sink).toHaveAttribute('aria-label', 'B2, 100');
		const id = await cell.getAttribute('id');
		await expect(sink).toHaveAttribute('aria-activedescendant', id!);
	});

test('the grid cells follow the window when it scrolls', async ({ page }) => {
	await newWorkbook(page);
	const table = editor(page).getByRole('grid', { name: 'Spreadsheet grid' });
	const lastRow = async () =>
		Math.max(
			...(await table
				.locator('[role="row"]')
				.evaluateAll((rows) => rows.map((row) => Number(row.getAttribute('aria-rowindex'))))),
		);
	const before = await lastRow();
	await expect(table.locator('[role="row"][aria-rowindex="2"]')).toHaveCount(1);
	// Scroll the grid itself, with no selection change (as the scroll bar or a wheel does).
	await grid(page)
		.locator('.xg-scroller')
		.evaluate((scroller) => {
			scroller.scrollTop = scroller.scrollHeight;
		});
	await expect.poll(lastRow).toBeGreaterThan(before + 10);
	// Row 2 scrolled away; row 1 stays, since it holds the active cell.
	await expect(table.locator('[role="row"][aria-rowindex="2"]')).toHaveCount(0);
	await expect(table.locator('[role="row"][aria-rowindex="1"]')).toHaveCount(1);
});
