import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { expect, test, type Page } from '@playwright/test';
import { editor, goToCell, newWorkbook, ribbon } from './helpers';

const fixture = (name: string) => [
	...readFileSync(
		resolve(process.cwd(), `../../src/core/xlsx/__fixtures__/excel-pivot-${name}.xlsx`),
	),
];

const open = async (page: Page, name: string) => {
	await newWorkbook(page);
	await editor(page).evaluate(
		async (node, data) =>
			(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
				new Uint8Array(data),
				'fields.xlsx',
			),
		fixture(name),
	);
};

/** A block of the first sheet: `rows` lines from `row`, `cols` cells from `col`. */
const block = (page: Page, row: number, col: number, rows: number, cols: number) =>
	editor(page).evaluate(
		(node, [row0, col0, height, width]) => {
			const sheet = (
				node as unknown as {
					workbook: { sheets: { rows: Map<number, Map<number, { value: unknown }>> }[] };
				}
			).workbook.sheets[0]!;
			return Array.from({ length: height! }, (_v, r) =>
				Array.from(
					{ length: width! },
					(_w, c) => sheet.rows.get(row0! + r)?.get(col0! + c)?.value ?? null,
				),
			);
		},
		[row, col, rows, cols],
	);

const fieldSettings = async (page: Page, cell: string) => {
	await goToCell(page, cell);
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await ribbon(page).getByRole('button', { name: 'PivotTable', exact: true }).last().click();
	await editor(page).getByRole('menuitem', { name: 'Field Settings', exact: true }).click();
	return editor(page).locator('[data-dialog="pivot-field-settings"]');
};
const reopen = async (page: Page) => {
	const bytes = await editor(page).evaluate(async (node) => [
		...new Uint8Array(
			await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
		),
	]);
	await editor(page).evaluate(
		async (node, data) =>
			(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
				new Uint8Array(data),
				'reopened.xlsx',
			),
		bytes,
	);
};

test('sets custom subtotals on a column field and reopens the report', async ({ page }) => {
	await open(page, 'colsubs-none-compact');
	// The fixture hides the Category subtotals, so the groups have no subtotal columns.
	expect((await block(page, 3, 1, 1, 6))[0]).toEqual([
		'Fruit',
		null,
		'Veg',
		null,
		'Grand Total',
		null,
	]);
	const dialog = await fieldSettings(page, 'B5');
	await expect(dialog.getByLabel('Field:', { exact: true }).locator('option')).toHaveText([
		'Category',
		'Year',
	]);
	await expect(dialog.locator('input[type=radio][value=none]')).toBeChecked();
	await dialog.locator('input[type=radio][value=custom]').check();
	await dialog.getByLabel('Sum', { exact: true }).check();
	await dialog.getByLabel('Average', { exact: true }).check();
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	const headers = [
		'Fruit',
		null,
		'Fruit Sum',
		'Fruit Average',
		'Veg',
		null,
		'Veg Sum',
		'Veg Average',
		'Grand Total',
	];
	expect((await block(page, 3, 1, 1, 9))[0]).toEqual(headers);
	// East: 2025 and 2026 detail, then the sum and the average of the Fruit group.
	expect((await block(page, 5, 0, 1, 5))[0]).toEqual(['East', 10, 10.5, 20.5, 6.833333333333333]);
	await reopen(page);
	expect((await block(page, 3, 1, 1, 9))[0]).toEqual(headers);
	const again = await fieldSettings(page, 'B5');
	await expect(again.locator('input[type=radio][value=custom]')).toBeChecked();
	await expect(again.getByLabel('Average', { exact: true })).toBeChecked();
	await again.getByRole('button', { name: 'Cancel', exact: true }).click();
});

test('sets custom subtotals on a row field of a Values-on-rows report', async ({ page }) => {
	await open(page, 'rv3-last-compact-top');
	const labels = async () => (await block(page, 2, 0, 44, 1)).map((line) => line[0]);
	expect(await labels()).toContain('Fruit Data');
	const dialog = await fieldSettings(page, 'A5');
	await expect(dialog.getByLabel('Field:', { exact: true }).locator('option')).toHaveText([
		'Category',
		'Region',
		'Product',
	]);
	await dialog.locator('input[type=radio][value=custom]').check();
	await dialog.getByLabel('Sum', { exact: true }).check();
	await dialog.getByLabel('Max', { exact: true }).check();
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	const after = await labels();
	// Below the Values, each function becomes one row per measure.
	expect(after.filter((label) => String(label).startsWith('Fruit '))).toEqual([
		'Fruit Sum of Amount',
		'Fruit Sum of Amount',
		'Fruit Max of Amount',
		'Fruit Max of Amount',
	]);
	await reopen(page);
	expect((await labels()).filter((label) => String(label).startsWith('Fruit '))).toEqual([
		'Fruit Sum of Amount',
		'Fruit Sum of Amount',
		'Fruit Max of Amount',
		'Fruit Max of Amount',
	]);
});

test('moves the Values among three row fields in the PivotTable settings', async ({ page }) => {
	await open(page, 'rv3-last-compact-top');
	await goToCell(page, 'A5');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await ribbon(page).getByRole('button', { name: 'PivotTable', exact: true }).last().click();
	await editor(page).getByRole('menuitem', { name: 'PivotTable Settings', exact: true }).click();
	const dialog = editor(page).locator('[data-dialog="pivot-settings"]');
	const placement = dialog.getByLabel('Values Placement:', { exact: true });
	await expect(placement.locator('option')).toHaveText([
		'Before row fields',
		'After Category',
		'After Region',
		'After row fields',
	]);
	await placement.selectOption({ label: 'After Category' });
	await expect(dialog).toContainText('Values are listed after Category.');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	// Category, then the measures under each category, then the regions.
	expect((await block(page, 2, 0, 6, 1)).map((line) => line[0])).toEqual([
		'Row Labels',
		'Fruit',
		'Data',
		'East',
		'Apple',
		'Pear',
	]);
	await reopen(page);
	expect((await block(page, 2, 0, 3, 1)).map((line) => line[0])).toEqual([
		'Row Labels',
		'Fruit',
		'Data',
	]);
});
