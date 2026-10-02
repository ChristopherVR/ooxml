import { expect, test } from '@playwright/test';
import {
	editor,
	formulaBarValue,
	goToCell,
	grid,
	newWorkbook,
	openSample,
	pageErrors,
	ribbon,
	sheetTabs,
	typeInActiveCell,
} from './helpers';

/** A cell's value from the live model (row and col zero-based). */
const cellValue = (
	page: import('@playwright/test').Page,
	sheet: number,
	row: number,
	col: number,
) =>
	editor(page).evaluate(
		(node, [s, r, c]) =>
			(
				node as unknown as {
					workbook: { sheets: { rows: Map<number, Map<number, { value: unknown }>> }[] };
				}
			).workbook.sheets[s!]!.rows.get(r!)?.get(c!)?.value ?? null,
		[sheet, row, col],
	);

test('formulas recalculate when their inputs change, in the sample too', async ({ page }) => {
	const errors = pageErrors(page);
	await openSample(page);
	await goToCell(page, 'B4');
	await typeInActiveCell(page, '100000');
	// N4 is =SUM(B4:M4): 624,640 - 42,100 + 100,000.
	await expect(grid(page)).toContainText('$682,540');
	await goToCell(page, 'N4');
	expect(await formulaBarValue(page)).toBe('=SUM(B4:M4)');
	expect(errors).toEqual([]);
});

test('undo and redo walk back and forth through edits', async ({ page }) => {
	await newWorkbook(page);
	await typeInActiveCell(page, 'first');
	await goToCell(page, 'A1');
	await typeInActiveCell(page, 'second');
	expect(await cellValue(page, 0, 0, 0)).toBe('second');
	await page.keyboard.press('Control+Z');
	expect(await cellValue(page, 0, 0, 0)).toBe('first');
	await page.keyboard.press('Control+Z');
	expect(await cellValue(page, 0, 0, 0)).toBeNull();
	await page.keyboard.press('Control+Y');
	expect(await cellValue(page, 0, 0, 0)).toBe('first');
});

test('rows insert and delete from the ribbon, moving formulas', async ({ page }) => {
	await newWorkbook(page);
	await typeInActiveCell(page, '5');
	await typeInActiveCell(page, '=A1*2');
	await editor(page).evaluate((node) =>
		(node as unknown as { select(ref: string): void }).select('1:1'),
	);
	await ribbon(page).locator('[data-command="cells.insert-split"]').first().click();
	expect(await cellValue(page, 0, 1, 0)).toBe(5);
	expect(await cellValue(page, 0, 2, 0)).toBe(10);
	await goToCell(page, 'A3');
	expect(await formulaBarValue(page)).toBe('=A2*2');
	await editor(page).evaluate((node) =>
		(node as unknown as { select(ref: string): void }).select('1:1'),
	);
	await ribbon(page).locator('[data-command="cells.delete-split"]').first().click();
	expect(await cellValue(page, 0, 0, 0)).toBe(5);
	expect(await cellValue(page, 0, 1, 0)).toBe(10);
});

test('sheet tabs add a sheet and rename it by double-click', async ({ page }) => {
	await newWorkbook(page);
	await sheetTabs(page).getByRole('button', { name: 'New sheet' }).click();
	await expect(sheetTabs(page)).toContainText('Sheet2');
	await sheetTabs(page).getByRole('tab', { name: 'Sheet2' }).dblclick();
	const field = sheetTabs(page).locator('input');
	await field.fill('Totals');
	await field.press('Enter');
	await expect(sheetTabs(page).getByRole('tab', { name: 'Totals' })).toBeVisible();
	const names = await editor(page).evaluate((node) =>
		(node as unknown as { workbook: { sheets: { name: string }[] } }).workbook.sheets.map(
			(s) => s.name,
		),
	);
	expect(names).toEqual(['Sheet1', 'Totals']);
});

test('read-only mode keeps the workbook unchanged and disables editing commands', async ({
	page,
}) => {
	await openSample(page);
	await editor(page).evaluate((node) => node.setAttribute('read-only', ''));
	await expect(ribbon(page).locator('[data-command="home.bold"]').first()).toBeDisabled();
	await goToCell(page, 'A20');
	await typeInActiveCell(page, 'blocked');
	expect(await cellValue(page, 0, 19, 0)).toBeNull();
	expect(await editor(page).evaluate((node) => (node as unknown as { dirty: boolean }).dirty)).toBe(
		false,
	);
});

test('clicking a chart selects it: Chart Design appears and Delete removes it', async ({
	page,
}) => {
	await openSample(page);
	const chartDesign = ribbon(page).getByRole('tab', { name: 'Chart Design' });
	await expect(chartDesign).toBeHidden();
	await grid(page)
		.locator('.xg-chart')
		.first()
		.click({ position: { x: 200, y: 60 } });
	await expect(chartDesign).toBeVisible();
	const count = () =>
		editor(page).evaluate(
			(node) =>
				(node as unknown as { workbook: { sheets: { drawings: unknown[] }[] } }).workbook.sheets[0]!
					.drawings.length,
		);
	expect(await count()).toBe(1);
	await page.keyboard.press('Delete');
	expect(await count()).toBe(0);
	await expect(chartDesign).toBeHidden();
	await page.keyboard.press('Control+Z');
	expect(await count()).toBe(1);
});
