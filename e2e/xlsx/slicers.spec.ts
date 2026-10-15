import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { expect, test, type Page } from '@playwright/test';
import { editor, goToCell, newWorkbook, ribbon } from './helpers';

const fixture = (name: string) => [
	...readFileSync(resolve(process.cwd(), `../../src/core/xlsx/__fixtures__/${name}`)),
];

async function open(page: Page, name: string) {
	await newWorkbook(page);
	await editor(page).evaluate(
		async (node, data) =>
			(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
				new Uint8Array(data),
				'slicers.xlsx',
			),
		fixture(name),
	);
}

const slicer = (page: Page, name: string) =>
	editor(page).locator(`.xg-slicer[aria-label="${name}"]`);
const item = (page: Page, slicerName: string, label: string) =>
	slicer(page, slicerName).locator('.xg-sl-item', { hasText: label });

/** Row 1-based data rows that the table currently hides. */
const hiddenRows = (page: Page) =>
	editor(page).evaluate((node) => {
		const sheet = (
			node as unknown as {
				workbook: { sheets: { rowInfo: Map<number, { hidden?: boolean }> }[] };
			}
		).workbook.sheets[0]!;
		return [...sheet.rowInfo].filter(([, info]) => info.hidden).map(([row]) => row + 1);
	});

test('a table slicer filters rows with click, Ctrl, Multi-Select and Clear Filter', async ({
	page,
}, info) => {
	await open(page, 'excel-slicer-table.xlsx');
	await expect(slicer(page, 'Category')).toBeVisible();
	await expect(slicer(page, 'Region')).toBeVisible();
	await expect(item(page, 'Category', 'Veg')).toHaveClass(/xg-sl-unselected/);
	expect(await hiddenRows(page)).toEqual([4, 5]);
	await item(page, 'Category', 'Veg').click();
	await expect(item(page, 'Category', 'Veg')).toHaveClass(/xg-sl-selected/);
	await expect(item(page, 'Category', 'Fruit')).toHaveClass(/xg-sl-unselected/);
	expect(await hiddenRows(page)).toEqual([2, 3, 6, 7]);
	await item(page, 'Category', 'Nuts').click({ modifiers: ['Control'] });
	expect(await hiddenRows(page)).toEqual([2, 3, 6]);
	await slicer(page, 'Category').locator('[data-slicer-action=multi]').click();
	await item(page, 'Category', 'Fruit').click();
	expect(await hiddenRows(page)).toEqual([]);
	await item(page, 'Category', 'Veg').click();
	await expect(slicer(page, 'Category').locator('[data-slicer-action=clear]')).toBeEnabled();
	await slicer(page, 'Category').locator('[data-slicer-action=clear]').click();
	expect(await hiddenRows(page)).toEqual([]);
	await page.screenshot({ path: info.outputPath('table-slicer.png') });
	// Filter again (Multi-Select off), save, reopen: the slicers and the filter come back.
	await slicer(page, 'Category').locator('[data-slicer-action=multi]').click();
	await item(page, 'Category', 'Nuts').click();
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
	await expect(slicer(page, 'Region')).toBeVisible();
	await expect(item(page, 'Category', 'Nuts')).toHaveClass(/xg-sl-selected/);
	await expect(item(page, 'Category', 'Fruit')).toHaveClass(/xg-sl-unselected/);
});

test('a PivotTable slicer filters the report and the Slicer tab styles it', async ({
	page,
}, info) => {
	await open(page, 'excel-slicer-pivot.xlsx');
	await expect(slicer(page, 'Category')).toBeVisible();
	// Fruit and Nuts are shown; Ctrl+click Veg shows it too, which clears the item filter.
	await item(page, 'Category', 'Veg').click({ modifiers: ['Control'] });
	const state = () =>
		editor(page).evaluate((node) => {
			const sheet = (
				node as unknown as {
					workbook: {
						sheets: {
							pivotTables?: { itemFilters?: unknown }[];
							rows: Map<number, Map<number, { value: unknown }>>;
						}[];
					};
				}
			).workbook.sheets[0]!;
			return {
				filters: sheet.pivotTables![0]!.itemFilters,
				labels: [2, 3, 4, 5, 6].map((row) => sheet.rows.get(row)?.get(0)?.value ?? null),
			};
		});
	expect((await state()).filters).toBeUndefined();
	expect((await state()).labels).toContain('Veg');
	await item(page, 'Category', 'Fruit').click();
	expect((await state()).labels).not.toContain('Veg');
	expect((await state()).labels).toContain('Fruit');
	await slicer(page, 'Category').locator('.xg-sl-head').click();
	await expect(ribbon(page).getByRole('tab', { name: 'Slicer', exact: true })).toBeVisible();
	await ribbon(page).getByRole('tab', { name: 'Slicer', exact: true }).click();
	await expect(ribbon(page).locator('[data-command="slicer.settings"]')).toBeVisible();
	await page.screenshot({ path: info.outputPath('pivot-slicer.png') });
});

test('Insert > Slicer adds a slicer for a table column and the Slicer tab formats it', async ({
	page,
}, info) => {
	await open(page, 'excel-slicer-table.xlsx');
	await goToCell(page, 'B3');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await ribbon(page).locator('[data-command="insert.slicer"]').click();
	const dialog = editor(page).locator('[data-dialog="insert-slicer"]');
	await dialog.getByLabel('Amount', { exact: true }).check();
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(slicer(page, 'Amount')).toBeVisible();
	await expect(item(page, 'Amount', '20')).toHaveClass(/xg-sl-selected/);
	await expect(ribbon(page).getByRole('tab', { name: 'Slicer', exact: true })).toBeVisible();
	await ribbon(page).getByRole('tab', { name: 'Slicer', exact: true }).click();
	await ribbon(page).locator('[data-command="slicer.settings"]').click();
	const settings = editor(page).locator('[data-dialog="slicer-settings"]');
	await settings.getByLabel('Caption:', { exact: true }).fill('Money');
	await settings.getByLabel('Sort items:', { exact: true }).selectOption('descending');
	await settings.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(slicer(page, 'Money').locator('.xg-sl-caption')).toHaveText('Money');
	await expect(slicer(page, 'Money').locator('.xg-sl-item').first()).toHaveText('20');
	await editor(page).evaluate((node) => (node as unknown as { undo(): void }).undo());
	await expect(slicer(page, 'Amount').locator('.xg-sl-caption')).toHaveText('Amount');
	await page.screenshot({ path: info.outputPath('insert-slicer.png') });
});
