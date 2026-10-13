import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { expect, test } from '@playwright/test';
import {
	editor,
	formulaBarValue,
	goToCell,
	newWorkbook,
	ribbon,
	typeInActiveCell,
} from './helpers';

test('authors Values on rows, refreshes, undoes and reopens axis settings', async ({
	page,
}, info) => {
	await newWorkbook(page);
	const rows = [
		['Category', 'Amount', 'Quantity', 'Year'],
		['Fruit', '10', '2', '2025'],
		['Fruit', '20', '3', '2026'],
		['Veg', '5', '4', '2025'],
	];
	for (const [r, line] of rows.entries())
		for (const [c, value] of line.entries()) {
			await goToCell(page, `${String.fromCharCode(65 + c)}${r + 1}`);
			await typeInActiveCell(page, value);
		}
	await goToCell(page, 'F1');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await ribbon(page).locator('[data-command="insert.pivot-table"]').click();
	const create = editor(page).locator('[data-dialog="create-pivot-table"]');
	await create.getByLabel('Source Range:', { exact: true }).fill('A1:D4');
	await create.getByLabel('Destination Cell:', { exact: true }).fill('F1');
	await create.getByLabel('Value Field:', { exact: true }).selectOption('Amount');
	await create.getByLabel('Column Field:', { exact: true }).selectOption('Year');
	await create.getByRole('button', { name: 'Add Value Field', exact: true }).click();
	const extra = create.locator('[data-pivot-value]');
	await extra.getByLabel('Value Field:', { exact: true }).selectOption('Quantity');
	await extra.getByLabel('Summarize Values By:', { exact: true }).selectOption('count');
	await create.getByLabel('Values Axis:', { exact: true }).selectOption('rows');
	await expect(create.getByLabel('Report Layout:', { exact: true })).toHaveValue('tabular');
	await expect(create.getByLabel('Second Row Field:', { exact: true })).toBeDisabled();
	await create.getByLabel('Repeat row labels', { exact: true }).check();
	await page.screenshot({ path: info.outputPath('pivot-values-on-rows.png') });
	await create.getByRole('button', { name: 'OK', exact: true }).click();
	await goToCell(page, 'J7');
	await expect.poll(() => formulaBarValue(page)).toBe('35');
	await goToCell(page, 'J8');
	await expect.poll(() => formulaBarValue(page)).toBe('3');
	await goToCell(page, 'B2');
	await typeInActiveCell(page, '100');
	await goToCell(page, 'F1');
	const menu = () =>
		ribbon(page).getByRole('button', { name: 'PivotTable', exact: true }).last().click();
	await menu();
	await editor(page).getByRole('menuitem', { name: 'Refresh PivotTable', exact: true }).click();
	await goToCell(page, 'J7');
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
				'row-values.xlsx',
			),
		bytes,
	);
	await goToCell(page, 'J7');
	await expect.poll(() => formulaBarValue(page)).toBe('125');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await menu();
	await editor(page).getByRole('menuitem', { name: 'PivotTable Settings', exact: true }).click();
	const settings = editor(page).locator('[data-dialog="pivot-settings"]');
	await expect(settings.getByLabel('Values Axis:', { exact: true })).toHaveValue('rows');
	await expect(settings.getByLabel('Repeat row labels', { exact: true })).toBeChecked();
	await settings.getByLabel('Values Axis:', { exact: true }).selectOption('columns');
	await settings.getByRole('button', { name: 'OK', exact: true }).click();
	await goToCell(page, 'K6');
	await expect.poll(() => formulaBarValue(page)).toBe('125');
	await page.keyboard.press('Control+Z');
	await goToCell(page, 'J7');
	await expect.poll(() => formulaBarValue(page)).toBe('125');
});

test('keeps native Values-on-rows definitions verbatim when accepting unchanged settings', async ({
	page,
}) => {
	const bytes = readFileSync(
		new URL(
			'../../src/core/xlsx/__fixtures__/excel-pivot-row-values-columns.xlsx',
			import.meta.url,
		),
	);
	const original = await JSZip.loadAsync(bytes);
	await newWorkbook(page);
	await editor(page).evaluate(
		async (node, data) =>
			(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
				new Uint8Array(data),
				'native-row-values.xlsx',
			),
		[...bytes],
	);
	await goToCell(page, 'E9');
	await expect.poll(() => formulaBarValue(page)).toBe('57');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await ribbon(page).getByRole('button', { name: 'PivotTable', exact: true }).last().click();
	await editor(page).getByRole('menuitem', { name: 'PivotTable Settings', exact: true }).click();
	const settings = editor(page).locator('[data-dialog="pivot-settings"]');
	await expect(settings.getByLabel('Values Axis:', { exact: true })).toHaveValue('rows');
	await expect(settings.getByLabel('Report Layout:', { exact: true })).toHaveValue('tabular');
	await expect(settings.locator('[data-pivot-value]')).toHaveCount(2);
	await settings.getByRole('button', { name: 'OK', exact: true }).click();
	const output = await editor(page).evaluate(async (node) => [
		...new Uint8Array(
			await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
		),
	]);
	const saved = await JSZip.loadAsync(new Uint8Array(output));
	for (const part of Object.keys(original.files).filter((part) => part.startsWith('xl/pivot')))
		expect(await saved.file(part)?.async('string')).toBe(
			await original.file(part)!.async('string'),
		);
});
