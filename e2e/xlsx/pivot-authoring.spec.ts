import { expect, test } from '@playwright/test';
import {
	editor,
	formulaBarValue,
	goToCell,
	newWorkbook,
	ribbon,
	typeInActiveCell,
} from './helpers';

test('create, configure, refresh, delete and undo PivotTables through the ribbon', async ({
	page,
}, info) => {
	await newWorkbook(page);
	for (const [address, value] of [
		['A1', 'Category'],
		['B1', 'Amount'],
		['A2', 'Fruit'],
		['B2', '10'],
		['A3', 'Fruit'],
		['B3', '20'],
		['A4', 'Veg'],
		['B4', '5'],
	]) {
		await goToCell(page, address!);
		await typeInActiveCell(page, value!);
	}
	await goToCell(page, 'D1');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await ribbon(page).locator('[data-command="insert.pivot-table"]').click();
	const create = editor(page).locator('[data-dialog="create-pivot-table"]');
	await create.getByLabel('Source Range:', { exact: true }).fill('A1:B4');
	await create.getByLabel('Destination Cell:', { exact: true }).fill('D1');
	await create.getByLabel('Row Field:', { exact: true }).selectOption('Category');
	await create.getByLabel('Value Field:', { exact: true }).selectOption('Amount');
	await create.getByRole('button', { name: 'OK', exact: true }).click();
	await goToCell(page, 'E4');
	await expect.poll(() => formulaBarValue(page)).toBe('35');
	await goToCell(page, 'B2');
	await typeInActiveCell(page, '100');
	await goToCell(page, 'D1');
	const menu = () =>
		ribbon(page).getByRole('button', { name: 'PivotTable', exact: true }).last().click();
	await menu();
	await editor(page).getByRole('menuitem', { name: 'Refresh PivotTable', exact: true }).click();
	await goToCell(page, 'E4');
	await expect.poll(() => formulaBarValue(page)).toBe('125');
	await menu();
	await editor(page).getByRole('menuitem', { name: 'PivotTable Settings', exact: true }).click();
	const settings = editor(page).locator('[data-dialog="pivot-settings"]');
	await settings.getByLabel('Summarize Values By:', { exact: true }).selectOption('count');
	await settings.getByRole('button', { name: 'OK', exact: true }).click();
	await expect.poll(() => formulaBarValue(page)).toBe('3');
	await page.screenshot({ path: info.outputPath('pivot-table.png') });
	const bytes = await editor(page).evaluate(async (node) => [
		...new Uint8Array(
			await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
		),
	]);
	await editor(page).evaluate(
		async (node, data) =>
			(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
				new Uint8Array(data),
				'pivot-roundtrip.xlsx',
			),
		bytes,
	);
	await goToCell(page, 'E4');
	await expect.poll(() => formulaBarValue(page)).toBe('3');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await menu();
	await editor(page).getByRole('menuitem', { name: 'Delete PivotTable', exact: true }).click();
	await expect.poll(() => formulaBarValue(page)).toBe('');
	await page.keyboard.press('Control+Z');
	await expect.poll(() => formulaBarValue(page)).toBe('3');
	await page.keyboard.press('Control+Y');
	await expect.poll(() => formulaBarValue(page)).toBe('');
});
