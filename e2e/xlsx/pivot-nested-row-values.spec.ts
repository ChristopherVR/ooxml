import { expect, test } from '@playwright/test';
import {
	editor,
	formulaBarValue,
	goToCell,
	newWorkbook,
	ribbon,
	typeInActiveCell,
} from './helpers';

for (const columns of [false, true])
	test(`authors tabular nested row Values with columns ${columns}, history and save/reopen`, async ({
		page,
	}) => {
		await newWorkbook(page);
		const rows = [
			['Category', 'Region', 'Amount', 'Quantity', 'Year'],
			['Fruit', 'East', '10', '2', '2025'],
			['Fruit', 'West', '20', '3', '2026'],
			['Veg', 'East', '5', '4', '2025'],
		];
		for (const [r, line] of rows.entries())
			for (const [c, value] of line.entries()) {
				await goToCell(page, `${String.fromCharCode(65 + c)}${r + 1}`);
				await typeInActiveCell(page, value);
			}
		await goToCell(page, 'G1');
		await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
		await ribbon(page).locator('[data-command="insert.pivot-table"]').click();
		const dialog = editor(page).locator('[data-dialog="create-pivot-table"]');
		await dialog.getByLabel('Source Range:', { exact: true }).fill('A1:E4');
		await dialog.getByLabel('Destination Cell:', { exact: true }).fill('G1');
		await dialog.getByLabel('Value Field:', { exact: true }).selectOption('Amount');
		await dialog.getByLabel('Second Row Field:', { exact: true }).selectOption('Region');
		await dialog.getByLabel('Report Layout:', { exact: true }).selectOption('tabular');
		if (columns) await dialog.getByLabel('Column Field:', { exact: true }).selectOption('Year');
		await dialog.getByRole('button', { name: 'Add Value Field', exact: true }).click();
		const extra = dialog.locator('[data-pivot-value]');
		await extra.getByLabel('Value Field:', { exact: true }).selectOption('Quantity');
		await extra.getByLabel('Summarize Values By:', { exact: true }).selectOption('count');
		await dialog.getByLabel('Values Axis:', { exact: true }).selectOption('rows');
		await expect(dialog.getByLabel('Second Row Field:', { exact: true })).toBeEnabled();
		await expect(
			dialog.getByLabel('Report Layout:', { exact: true }).locator('option[value="compact"]'),
		).toHaveAttribute('disabled', '');
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		const total = columns ? 'L13' : 'J12';
		await goToCell(page, total);
		await expect.poll(() => formulaBarValue(page)).toBe('35');
		await goToCell(page, 'C2');
		await typeInActiveCell(page, '100');
		await goToCell(page, 'G1');
		const menu = () =>
			ribbon(page).getByRole('button', { name: 'PivotTable', exact: true }).last().click();
		await menu();
		await editor(page).getByRole('menuitem', { name: 'Refresh PivotTable', exact: true }).click();
		await goToCell(page, total);
		await expect.poll(() => formulaBarValue(page)).toBe('125');
		await page.keyboard.press('Control+Z');
		await expect.poll(() => formulaBarValue(page)).toBe('35');
		await page.keyboard.press('Control+Y');
		await expect.poll(() => formulaBarValue(page)).toBe('125');
		const bytes = await editor(page).evaluate(async (node) => [
			...new Uint8Array(
				await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
			),
		]);
		await editor(page).evaluate(
			async (node, data) =>
				(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
					new Uint8Array(data),
					'nested-values.xlsx',
				),
			bytes,
		);
		await goToCell(page, total);
		await expect.poll(() => formulaBarValue(page)).toBe('125');
		await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
		await menu();
		await editor(page).getByRole('menuitem', { name: 'PivotTable Settings', exact: true }).click();
		const settings = editor(page).locator('[data-dialog="pivot-settings"]');
		await expect(settings.getByLabel('Second Row Field:', { exact: true })).toHaveValue('Region');
		await expect(settings.getByLabel('Values Axis:', { exact: true })).toHaveValue('rows');
		await settings.getByLabel('Show subtotals', { exact: true }).uncheck();
		await settings.getByRole('button', { name: 'OK', exact: true }).click();
		await goToCell(page, columns ? 'L9' : 'J8');
		await expect.poll(() => formulaBarValue(page)).toBe('125');
	});
