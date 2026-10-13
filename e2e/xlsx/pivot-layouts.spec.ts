import { expect, test } from '@playwright/test';
import {
	editor,
	formulaBarValue,
	goToCell,
	newWorkbook,
	ribbon,
	typeInActiveCell,
} from './helpers';

test('authors compact reports, switches outline subtotals and reopens saved settings', async ({
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
	for (const [r, row] of rows.entries())
		for (const [c, value] of row.entries()) {
			await goToCell(page, `${String.fromCharCode(65 + c)}${r + 1}`);
			await typeInActiveCell(page, value);
		}
	await goToCell(page, 'E1');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await ribbon(page).locator('[data-command="insert.pivot-table"]').click();
	const create = editor(page).locator('[data-dialog="create-pivot-table"]');
	await create.getByLabel('Source Range:', { exact: true }).fill('A1:C6');
	await create.getByLabel('Destination Cell:', { exact: true }).fill('E1');
	await expect(create.getByLabel('Report Layout:', { exact: true })).toHaveValue('compact');
	await create.getByLabel('Second Row Field:', { exact: true }).selectOption('Region');
	await expect(create.getByLabel('Repeat row labels', { exact: true })).toBeDisabled();
	await create.getByRole('button', { name: 'OK', exact: true }).click();
	await goToCell(page, 'F2');
	await expect.poll(() => formulaBarValue(page)).toBe('37');
	const settings = editor(page).locator('[data-dialog="pivot-settings"]');
	const openSettings = async () => {
		await ribbon(page).getByRole('button', { name: 'PivotTable', exact: true }).last().click();
		await editor(page).getByRole('menuitem', { name: 'PivotTable Settings', exact: true }).click();
	};
	await openSettings();
	await settings.getByLabel('Report Layout:', { exact: true }).selectOption('outline');
	await settings.getByLabel('Subtotal Position:', { exact: true }).selectOption('bottom');
	await expect(settings.getByLabel('Repeat row labels', { exact: true })).toBeEnabled();
	await page.screenshot({ path: info.outputPath('pivot-outline-settings.png') });
	await settings.getByRole('button', { name: 'OK', exact: true }).click();
	await goToCell(page, 'G5');
	await expect.poll(() => formulaBarValue(page)).toBe('37');
	await page.keyboard.press('Control+Z');
	await goToCell(page, 'F2');
	await expect.poll(() => formulaBarValue(page)).toBe('37');
	await page.keyboard.press('Control+Y');
	await goToCell(page, 'G10');
	await expect.poll(() => formulaBarValue(page)).toBe('57');
	const bytes = await editor(page).evaluate(async (node) => [
		...new Uint8Array(
			await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
		),
	]);
	await editor(page).evaluate(
		async (node, data) =>
			(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
				new Uint8Array(data),
				'pivot-layouts.xlsx',
			),
		bytes,
	);
	await goToCell(page, 'G10');
	await expect.poll(() => formulaBarValue(page)).toBe('57');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await openSettings();
	await expect(settings.getByLabel('Report Layout:', { exact: true })).toHaveValue('outline');
	await expect(settings.getByLabel('Subtotal Position:', { exact: true })).toHaveValue('bottom');
	await settings.getByRole('button', { name: 'Cancel', exact: true }).click();
});
