import { expect, test } from '@playwright/test';
import {
	editor,
	formulaBarValue,
	goToCell,
	newWorkbook,
	ribbon,
	typeInActiveCell,
} from './helpers';

test('authors hierarchical PivotTables, toggles subtotals, refreshes and reopens saved layout', async ({
	page,
}, info) => {
	await newWorkbook(page);
	const rows = [
		['Category', 'Region', 'Amount'],
		['Fruit', 'East', '10'],
		['Fruit', 'West', '20'],
		['Veg', 'East', '5'],
		['Veg', 'West', '15'],
		['Fruit', 'East', '7'],
	];
	for (const [row, line] of rows.entries())
		for (const [col, value] of line.entries()) {
			await goToCell(page, `${String.fromCharCode(65 + col)}${row + 1}`);
			await typeInActiveCell(page, value);
		}
	await goToCell(page, 'E1');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await ribbon(page).locator('[data-command="insert.pivot-table"]').click();
	const create = editor(page).locator('[data-dialog="create-pivot-table"]');
	await create.getByLabel('Source Range:', { exact: true }).fill('A1:C6');
	await create.getByLabel('Destination Cell:', { exact: true }).fill('E1');
	await create.getByLabel('Row Field:', { exact: true }).selectOption('Category');
	await create.getByLabel('Second Row Field:', { exact: true }).selectOption('Region');
	await create.getByLabel('Value Field:', { exact: true }).selectOption('Amount');
	await create.getByRole('button', { name: 'OK', exact: true }).click();
	await goToCell(page, 'G4');
	await expect.poll(() => formulaBarValue(page)).toBe('37');
	const menu = () =>
		ribbon(page).getByRole('button', { name: 'PivotTable', exact: true }).last().click();
	await menu();
	await editor(page).getByRole('menuitem', { name: 'PivotTable Settings', exact: true }).click();
	const settings = editor(page).locator('[data-dialog="pivot-settings"]');
	await settings.getByLabel('Show subtotals below each group', { exact: true }).uncheck();
	await settings.getByLabel('Repeat row labels', { exact: true }).check();
	await page.screenshot({ path: info.outputPath('pivot-row-layout.png') });
	await settings.getByRole('button', { name: 'OK', exact: true }).click();
	await goToCell(page, 'E3');
	await expect.poll(() => formulaBarValue(page)).toBe('Fruit');
	await page.keyboard.press('Control+Z');
	await goToCell(page, 'E4');
	await expect.poll(() => formulaBarValue(page)).toBe('Fruit Total');
	await page.keyboard.press('Control+Y');
	await goToCell(page, 'C2');
	await typeInActiveCell(page, '100');
	await goToCell(page, 'E1');
	await menu();
	await editor(page).getByRole('menuitem', { name: 'Refresh PivotTable', exact: true }).click();
	await goToCell(page, 'G6');
	await expect.poll(() => formulaBarValue(page)).toBe('147');
	const bytes = await editor(page).evaluate(async (node) => [
		...new Uint8Array(
			await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
		),
	]);
	await editor(page).evaluate(
		async (node, data) =>
			(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
				new Uint8Array(data),
				'pivot-rows.xlsx',
			),
		bytes,
	);
	await goToCell(page, 'G6');
	await expect.poll(() => formulaBarValue(page)).toBe('147');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await menu();
	await editor(page).getByRole('menuitem', { name: 'PivotTable Settings', exact: true }).click();
	await expect(settings.getByLabel('Second Row Field:', { exact: true })).toHaveValue('Region');
	await expect(
		settings.getByLabel('Show subtotals below each group', { exact: true }),
	).not.toBeChecked();
	await expect(settings.getByLabel('Repeat row labels', { exact: true })).toBeChecked();
	await settings.getByRole('button', { name: 'Cancel', exact: true }).click();
});
