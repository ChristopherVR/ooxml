import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { expect, test, type Page } from '@playwright/test';
import {
	clickRibbonControl,
	editor,
	formulaBarValue,
	goToCell,
	newWorkbook,
	ribbon,
	sheetTabs,
	typeInActiveCell,
} from './helpers';

async function loadPivotWorkbook(page: Page) {
	const fixture = readFileSync(
		resolve(process.cwd(), '../../src/core/xlsx/__fixtures__/excel-pivot-rows.xlsx'),
	);
	await editor(page).evaluate(
		async (node, bytes) =>
			(node as unknown as { load(data: Uint8Array, name: string): Promise<void> }).load(
				new Uint8Array(bytes),
				'pivot.xlsx',
			),
		[...fixture],
	);
}

test.beforeEach(async ({ page }) => newWorkbook(page));

test('opens PivotTable contextual controls and reopens a report changed through the ribbon', async ({
	page,
}) => {
	const book = editor(page);
	const tabs = ribbon(page);
	const renameControl = tabs.locator('[data-tab="pivot-analyze"] [data-command="pivot.name"]');
	await expect(renameControl).toHaveCount(0);
	await loadPivotWorkbook(page);
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
	const gallery = tabs.getByRole('button', { name: 'PivotTable Styles', exact: true });
	await gallery.press('Enter');
	await expect(
		book.getByRole('menu', { name: 'PivotTable Styles', exact: true }).getByRole('menuitem'),
	).toHaveCount(85);
	await book.getByRole('menuitem', { name: 'Medium Style 28', exact: true }).press('Enter');
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
		display: { style: { name: 'PivotStyleMedium28', showRowStripes: true } },
	});
});

test('keeps cached totals when styling a report and updates them on Refresh', async ({ page }) => {
	await loadPivotWorkbook(page);
	await sheetTabs(page).getByRole('tab', { name: 'Data', exact: true }).click();
	await goToCell(page, 'C2');
	await typeInActiveCell(page, '30');
	await sheetTabs(page).getByRole('tab', { name: 'Summary', exact: true }).click();
	await goToCell(page, 'C8');
	const tabs = ribbon(page);
	await tabs.getByRole('tab', { name: 'PivotTable Design' }).click();
	await tabs.getByRole('button', { name: 'Banded Rows', exact: true }).click();
	await expect.poll(() => formulaBarValue(page)).toBe('57');
	await tabs.getByRole('tab', { name: 'PivotTable Analyze' }).click();
	await tabs.getByRole('button', { name: 'Refresh PivotTable', exact: true }).click();
	await expect.poll(() => formulaBarValue(page)).toBe('77');
});

test('applies a wildcard through the Label Equals filter', async ({ page }) => {
	await loadPivotWorkbook(page);
	await goToCell(page, 'A2');
	const tabs = ribbon(page);
	await tabs.getByRole('tab', { name: 'PivotTable Analyze' }).click();
	await clickRibbonControl(page, tabs.locator('[data-command="pivot.condition"]'));
	const dialog = editor(page).locator('[data-dialog="pivot-condition"]');
	await dialog.getByLabel('Filter:', { exact: true }).selectOption('captionEqual');
	await dialog.getByLabel('Value:', { exact: true }).fill('f?u*');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await goToCell(page, 'C5');
	await expect.poll(() => formulaBarValue(page)).toBe('37');
});
