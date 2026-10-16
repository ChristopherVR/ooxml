import { expect, test } from '@playwright/test';
import {
	editor,
	formulaBarValue,
	goToCell,
	newWorkbook,
	ribbon,
	typeInActiveCell,
} from './helpers';

/** The text of a cell, read through the formula bar. */
async function cellText(page: Parameters<typeof newWorkbook>[0], ref: string): Promise<string> {
	await goToCell(page, ref);
	return formulaBarValue(page);
}

async function enterData(page: Parameters<typeof newWorkbook>[0], rows: string[][]) {
	for (const [r, line] of rows.entries())
		for (const [c, value] of line.entries()) {
			await goToCell(page, `${String.fromCharCode(65 + c)}${r + 1}`);
			await typeInActiveCell(page, value);
		}
}

async function openCreate(page: Parameters<typeof newWorkbook>[0], source: string) {
	await goToCell(page, 'F1');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await ribbon(page).locator('[data-command="insert.pivot-table"]').click();
	const dialog = editor(page).locator('[data-dialog="create-pivot-table"]');
	await dialog.getByLabel('Source Range:', { exact: true }).fill(source);
	await dialog.getByLabel('Destination Cell:', { exact: true }).fill('F1');
	return dialog;
}

test('authors Values on rows under a report filter, refreshes, undoes and switches the axis', async ({
	page,
}) => {
	await newWorkbook(page);
	await enterData(page, [
		['Category', 'Amount', 'Quantity', 'Year'],
		['Fruit', '10', '2', '2025'],
		['Fruit', '20', '3', '2026'],
		['Veg', '5', '4', '2025'],
	]);
	const dialog = await openCreate(page, 'A1:D4');
	await dialog.getByLabel('Value Field:', { exact: true }).selectOption('Amount');
	await dialog.getByRole('button', { name: 'Add Value Field', exact: true }).click();
	const extra = dialog.locator('[data-pivot-value]');
	await extra.getByLabel('Value Field:', { exact: true }).selectOption('Quantity');
	await extra.getByLabel('Summarize Values By:', { exact: true }).selectOption('count');
	await dialog.getByLabel('Values Axis:', { exact: true }).selectOption('rows');
	// Both choices stay available together.
	await dialog.getByLabel('Report Filter Field:', { exact: true }).selectOption('Year');
	await expect(dialog.getByLabel('Report Filter Field:', { exact: true })).toBeEnabled();
	await expect(dialog.getByLabel('Values Axis:', { exact: true })).toHaveValue('rows');
	await dialog.getByLabel('Report Filter Item:', { exact: true }).selectOption({ label: '2025' });
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();

	// F1:G1 hold the report filter; the report starts on row 3 and ends with Total rows.
	await expect.poll(() => cellText(page, 'F1')).toBe('Year');
	await expect.poll(() => cellText(page, 'G1')).toBe('2025');
	await expect.poll(() => cellText(page, 'F3')).toBe('Row Labels');
	await goToCell(page, 'G10');
	await expect.poll(() => formulaBarValue(page)).toBe('15');
	await goToCell(page, 'G11');
	await expect.poll(() => formulaBarValue(page)).toBe('2');

	await goToCell(page, 'B2');
	await typeInActiveCell(page, '100');
	await goToCell(page, 'F3');
	const menu = () =>
		ribbon(page).getByRole('button', { name: 'PivotTable', exact: true }).last().click();
	await menu();
	await editor(page).getByRole('menuitem', { name: 'Refresh PivotTable', exact: true }).click();
	await goToCell(page, 'G10');
	await expect.poll(() => formulaBarValue(page)).toBe('105');
	await page.keyboard.press('Control+Z');
	await expect.poll(() => formulaBarValue(page)).toBe('15');
	await page.keyboard.press('Control+Y');
	await expect.poll(() => formulaBarValue(page)).toBe('105');

	// Moving the values to the columns keeps the report filter.
	await goToCell(page, 'F3');
	await menu();
	await editor(page).getByRole('menuitem', { name: 'PivotTable Settings', exact: true }).click();
	const settings = editor(page).locator('[data-dialog="pivot-settings"]');
	await expect(settings.getByLabel('Report Filter Field:', { exact: true })).toHaveValue('Year');
	await settings.getByLabel('Values Axis:', { exact: true }).selectOption('columns');
	await settings.getByRole('button', { name: 'OK', exact: true }).click();
	await expect.poll(() => cellText(page, 'F1')).toBe('Year');
	await expect.poll(() => cellText(page, 'G1')).toBe('2025');
});

test('captions a new report in the language of the editor', async ({ page }) => {
	await newWorkbook(page);
	await enterData(page, [
		['Category', 'Amount'],
		['Fruit', '10'],
		['Veg', '5'],
	]);
	const dialog = await openCreate(page, 'A1:B3');
	await dialog.getByLabel('Value Field:', { exact: true }).selectOption('Amount');
	// Excel captions a report in its UI language: the editor's language decides here.
	await editor(page).evaluate((node) => node.setAttribute('locale', 'de'));
	await editor(page)
		.locator('[data-dialog="create-pivot-table"]')
		.getByRole('button', { name: 'OK', exact: true })
		.click();
	await expect.poll(() => cellText(page, 'F1')).toBe('Zeilenbeschriftungen');
	await expect.poll(() => cellText(page, 'G1')).toBe('Summe von Amount');
	await expect.poll(() => cellText(page, 'F4')).toBe('Gesamtergebnis');
});
