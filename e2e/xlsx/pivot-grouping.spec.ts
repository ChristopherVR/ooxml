import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { expect, test, type Page } from '@playwright/test';
import { editor, goToCell, newWorkbook, ribbon } from './helpers';

const fixture = () => [
	...readFileSync(
		resolve(process.cwd(), '../../src/core/xlsx/__fixtures__/excel-pivot-group-number.xlsx'),
	),
];

const labels = (page: Page) =>
	editor(page).evaluate((node) => {
		const sheet = (
			node as unknown as {
				workbook: { sheets: { rows: Map<number, Map<number, { value: unknown }>> }[] };
			}
		).workbook.sheets[0]!;
		return [3, 4, 5, 6, 7, 8].map((row) => sheet.rows.get(row)?.get(0)?.value ?? null);
	});

test('groups and ungroups a numeric PivotTable field, then saves and reopens it', async ({
	page,
}) => {
	await newWorkbook(page);
	await editor(page).evaluate(
		async (node, data) =>
			(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
				new Uint8Array(data),
				'group.xlsx',
			),
		fixture(),
	);
	expect(await labels(page)).toEqual(['0-9', '10-19', '20-29', '40-50', 'Grand Total', null]);
	await goToCell(page, 'A4');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	const menu = () =>
		ribbon(page).getByRole('button', { name: 'PivotTable', exact: true }).last().click();
	await menu();
	await editor(page).getByRole('menuitem', { name: 'Group Field', exact: true }).click();
	const dialog = editor(page).locator('[data-dialog="pivot-group"]');
	await expect(dialog.getByLabel('Starting at:', { exact: true })).toHaveValue('0');
	await dialog.getByLabel('Ending at:', { exact: true }).fill('60');
	await dialog.getByLabel('By:', { exact: true }).fill('30');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	expect(await labels(page)).toEqual(['0-29', '30-60', 'Grand Total', null, null, null]);
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
	expect(await labels(page)).toEqual(['0-29', '30-60', 'Grand Total', null, null, null]);
	await goToCell(page, 'A4');
	await menu();
	await editor(page).getByRole('menuitem', { name: 'Ungroup', exact: true }).click();
	expect(await labels(page)).toEqual([3, 12, 17, 25, 29, 44]);
});
