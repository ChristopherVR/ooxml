import { expect, test } from '@playwright/test';
import {
	editor,
	formulaBarValue,
	goToCell,
	newWorkbook,
	ribbon,
	typeInActiveCell,
} from './helpers';

test('authors three pivot measures with nested rows, refreshes, undoes and reopens saved settings', async ({
	page,
}, info) => {
	await newWorkbook(page);
	const rows = [
		['Category', 'Region', 'Amount', 'Quantity', 'Year'],
		['Fruit', 'East', '10', '2', '2025'],
		['Fruit', 'West', '20', '3', '2026'],
		['Veg', 'East', '5', '4', '2025'],
		['Veg', 'West', '15', '5', '2026'],
		['Fruit', 'East', '7', '6', '2026'],
	];
	for (const [r, line] of rows.entries())
		for (const [c, value] of line.entries()) {
			await goToCell(page, `${String.fromCharCode(65 + c)}${r + 1}`);
			await typeInActiveCell(page, value);
		}
	await goToCell(page, 'G1');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await ribbon(page).locator('[data-command="insert.pivot-table"]').click();
	const create = editor(page).locator('[data-dialog="create-pivot-table"]');
	await create.getByLabel('Source Range:', { exact: true }).fill('A1:E6');
	await create.getByLabel('Destination Cell:', { exact: true }).fill('G1');
	await create.getByLabel('Row Field:', { exact: true }).selectOption('Category');
	await create.getByLabel('Second Row Field:', { exact: true }).selectOption('Region');
	await create.getByLabel('Column Field:', { exact: true }).selectOption('Year');
	await create.getByLabel('Value Field:', { exact: true }).selectOption('Amount');
	await create.getByLabel('Report Layout:', { exact: true }).selectOption('tabular');
	await create.getByRole('button', { name: 'Add Value Field', exact: true }).click();
	const extras = create.locator('[data-pivot-value]');
	await extras.nth(0).getByLabel('Value Field:', { exact: true }).selectOption('Quantity');
	await extras.nth(0).getByLabel('Summarize Values By:', { exact: true }).selectOption('count');
	await create.getByRole('button', { name: 'Add Value Field', exact: true }).click();
	await extras.nth(1).getByLabel('Value Field:', { exact: true }).selectOption('Amount');
	await extras.nth(1).getByLabel('Summarize Values By:', { exact: true }).selectOption('average');
	await expect(create.getByLabel('Report Layout:', { exact: true })).toHaveValue('tabular');
	await page.screenshot({ path: info.outputPath('pivot-three-values.png') });
	await create.getByRole('button', { name: 'OK', exact: true }).click();
	await goToCell(page, 'O10');
	await expect.poll(() => formulaBarValue(page)).toBe('57');
	await goToCell(page, 'P10');
	await expect.poll(() => formulaBarValue(page)).toBe('5');
	await goToCell(page, 'Q10');
	await expect.poll(() => formulaBarValue(page)).toBe('11.4');
	await goToCell(page, 'C2');
	await typeInActiveCell(page, '100');
	await goToCell(page, 'G1');
	const menu = () =>
		ribbon(page).getByRole('button', { name: 'PivotTable', exact: true }).last().click();
	await menu();
	await editor(page).getByRole('menuitem', { name: 'Refresh PivotTable', exact: true }).click();
	await goToCell(page, 'O10');
	await expect.poll(() => formulaBarValue(page)).toBe('147');
	await page.keyboard.press('Control+Z');
	await expect.poll(() => formulaBarValue(page)).toBe('57');
	await page.keyboard.press('Control+Y');
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
				'pivot-values.xlsx',
			),
		bytes,
	);
	await goToCell(page, 'Q10');
	await expect.poll(() => formulaBarValue(page)).toBe('29.4');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await menu();
	await editor(page).getByRole('menuitem', { name: 'PivotTable Settings', exact: true }).click();
	const settings = editor(page).locator('[data-dialog="pivot-settings"]');
	await expect(settings.locator('[data-pivot-value]')).toHaveCount(2);
	await expect(
		settings.locator('[data-pivot-value]').nth(0).getByLabel('Value Field:', { exact: true }),
	).toHaveValue('Quantity');
	await expect(
		settings
			.locator('[data-pivot-value]')
			.nth(1)
			.getByLabel('Summarize Values By:', { exact: true }),
	).toHaveValue('average');
	await settings.getByRole('button', { name: 'Cancel', exact: true }).click();
});
