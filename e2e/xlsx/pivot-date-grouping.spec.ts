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
				'grouped.xlsx',
			),
		fixture(name),
	);
};

/** The first column of the Summary sheet from row 1 down, as the grid holds it. */
const column = (page: Page, count: number) =>
	editor(page).evaluate((node, rows) => {
		const sheet = (
			node as unknown as {
				workbook: { sheets: { rows: Map<number, Map<number, { value: unknown }>> }[] };
			}
		).workbook.sheets[0]!;
		return Array.from({ length: rows }, (_v, row) => sheet.rows.get(row)?.get(0)?.value ?? null);
	}, count);

const menu = (page: Page) =>
	ribbon(page).getByRole('button', { name: 'PivotTable', exact: true }).last().click();

test('regroups the dates of a report that has a report filter, then saves and reopens it', async ({
	page,
}) => {
	await open(page, 'group-page-west');
	expect(await column(page, 11)).toEqual([
		'Region',
		null,
		'Row Labels',
		'2023',
		'Dec',
		'2024',
		'Feb',
		'Apr',
		'Nov',
		'Grand Total',
		null,
	]);
	await goToCell(page, 'A4');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await menu(page);
	await editor(page).getByRole('menuitem', { name: 'Group Selection', exact: true }).click();
	const dialog = editor(page).locator('[data-dialog="pivot-group-dates"]');
	await expect(dialog.getByLabel('Months', { exact: true })).toBeChecked();
	await expect(dialog.getByLabel('Years', { exact: true })).toBeChecked();
	await expect(dialog.getByLabel('Quarters', { exact: true })).not.toBeChecked();
	await dialog.getByLabel('Quarters', { exact: true }).check();
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	const grouped = [
		'Region',
		null,
		'Row Labels',
		'2023',
		'Qtr4',
		'Dec',
		'2024',
		'Qtr1',
		'Feb',
		'Qtr2',
		'Apr',
		'Qtr4',
		'Nov',
		'Grand Total',
		null,
	];
	expect(await column(page, grouped.length)).toEqual(grouped);
	const bytes = await editor(page).evaluate(async (node) => [
		...new Uint8Array(
			await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
		),
	]);
	await editor(page).evaluate(
		async (node, data) =>
			(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
				new Uint8Array(data),
				'regrouped.xlsx',
			),
		bytes,
	);
	expect(await column(page, grouped.length)).toEqual(grouped);
	// The report filter survives: only the West rows are listed and the filter cell is intact.
	await goToCell(page, 'B1');
	await menu(page);
	await editor(page).getByRole('menuitem', { name: 'Ungroup', exact: true }).click();
	const ungrouped = await column(page, 4);
	expect(ungrouped.slice(0, 3)).toEqual(['Region', null, 'Row Labels']);
	expect(typeof ungrouped[3]).toBe('number');
});

test('keeps a timeline range when the dates are grouped differently', async ({ page }) => {
	await open(page, 'group-timeline');
	expect(await column(page, 8)).toEqual([
		null,
		null,
		'Row Labels',
		'2024',
		'Jan',
		'Feb',
		'Apr',
		'Grand Total',
	]);
	await goToCell(page, 'A4');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await menu(page);
	await editor(page).getByRole('menuitem', { name: 'Group Selection', exact: true }).click();
	const dialog = editor(page).locator('[data-dialog="pivot-group-dates"]');
	await dialog.getByLabel('Months', { exact: true }).uncheck();
	await dialog.getByLabel('Quarters', { exact: true }).check();
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	// Only the dates inside the timeline range remain, now in quarters.
	expect(await column(page, 6)).toEqual([null, null, 'Row Labels', '2024', 'Qtr1', 'Qtr2']);
});
