import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { expect, test } from '@playwright/test';
import { editor, goToCell, newWorkbook, ribbon } from './helpers';

test('opens PivotTable contextual controls and reopens a report changed through the ribbon', async ({
	page,
}) => {
	await newWorkbook(page);
	const book = editor(page);
	const tabs = ribbon(page);
	const fixture = readFileSync(
		resolve(process.cwd(), '../../src/core/xlsx/__fixtures__/excel-pivot-rows.xlsx'),
	);
	await book.evaluate(
		async (node, bytes) =>
			(node as unknown as { load(data: Uint8Array, name: string): Promise<void> }).load(
				new Uint8Array(bytes),
				'pivot.xlsx',
			),
		[...fixture],
	);
	const analyze = tabs.getByRole('tab', { name: 'PivotTable Analyze' });
	const design = tabs.getByRole('tab', { name: 'PivotTable Design' });
	await goToCell(page, 'J12');
	await expect(analyze).toHaveCount(0);
	await expect(design).toHaveCount(0);
	await goToCell(page, 'A2');
	await analyze.focus();
	await analyze.press('Enter');
	await tabs.getByRole('button', { name: 'PivotTable Name', exact: true }).click();
	const rename = book.locator('[data-dialog="pivot-name"]');
	await rename.getByRole('textbox', { name: 'PivotTable Name', exact: true }).fill('NewSales');
	await rename.getByRole('button', { name: 'OK', exact: true }).click();
	await design.click();
	await tabs.getByRole('button', { name: 'Report Layout', exact: true }).click();
	await book.getByRole('menuitemcheckbox', { name: 'Outline Form', exact: true }).click();
	const stripes = tabs.getByRole('button', { name: 'Banded Rows', exact: true });
	await stripes.focus();
	await stripes.press('Space');
	await expect(stripes).toHaveAttribute('aria-pressed', 'true');
	const result = await book.evaluate(async (node) => {
		const control = node as unknown as {
			save(): Promise<Blob>;
			load(data: Uint8Array, name: string): Promise<void>;
			workbook: { sheets: { pivotTables?: unknown[] }[] };
		};
		await control.load(new Uint8Array(await (await control.save()).arrayBuffer()), 'reopened.xlsx');
		return control.workbook.sheets[0]?.pivotTables?.[0];
	});
	expect(result).toMatchObject({
		name: 'NewSales',
		layout: 'outline',
		display: { style: { showRowStripes: true } },
	});
});
