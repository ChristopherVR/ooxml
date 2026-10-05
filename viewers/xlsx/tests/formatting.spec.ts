import { expect, test, type Page } from '@playwright/test';
import { editor, goToCell, grid, newWorkbook, ribbon, typeInActiveCell } from './helpers';

/** The resolved style of a cell (row and col zero-based) from the live model. */
const cellStyle = (page: Page, row: number, col: number) =>
	editor(page).evaluate(
		(node, [r, c]) => {
			const wb = (
				node as unknown as {
					workbook: {
						styles: {
							font: { bold?: boolean };
							fill: { type: string; pattern?: string };
							numFmt?: string;
						}[];
						sheets: { rows: Map<number, Map<number, { styleId?: number }>> }[];
					};
				}
			).workbook;
			return wb.styles[wb.sheets[0]!.rows.get(r!)?.get(c!)?.styleId ?? 0];
		},
		[row, col],
	);

test('Bold, Fill Color and the number format list format the selection', async ({ page }) => {
	await newWorkbook(page);
	await typeInActiveCell(page, '1234.5');
	await goToCell(page, 'A1');
	await ribbon(page).locator('[data-command="home.bold"]').first().click();
	expect((await cellStyle(page, 0, 0))?.font.bold).toBe(true);
	await ribbon(page).locator('[data-command="home.fill-color"]').first().click();
	expect((await cellStyle(page, 0, 0))?.fill).toMatchObject({ type: 'pattern', pattern: 'solid' });
	await ribbon(page)
		.locator('select[data-command="home.number-format"]')
		.selectOption({ label: 'Currency' });
	await expect.poll(async () => (await cellStyle(page, 0, 0))?.numFmt).toContain('$');
	// Display text depends on font metrics: without Calibri (Linux CI) bold `$1,234.50` is wider
	// than the default column, and the grid shows `#######` exactly as Excel does.
	await expect(grid(page)).toContainText(/\$1,234\.50|#{4,}/);
	await page.keyboard.press('Control+B');
	expect((await cellStyle(page, 0, 0))?.font.bold).toBeFalsy();
});
