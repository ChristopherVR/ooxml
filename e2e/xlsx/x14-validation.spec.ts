import { fileURLToPath } from 'node:url';
import { expect, test } from '@playwright/test';
import type { Workbook } from 'ooxml-core/xlsx';
import {
	clickRibbonControl,
	editor,
	FRAMEWORKS,
	goToCell,
	grid,
	landingFileInput,
	openLanding,
	pageErrors,
	ribbon,
} from './helpers';

interface ValidationEditor extends HTMLElement {
	workbook: Workbook;
	saveBytes(): Promise<Uint8Array>;
	load(bytes: Uint8Array, name: string): Promise<void>;
}
const fixture = fileURLToPath(
	new URL('../../src/core/xlsx/__fixtures__/excel-x14-validation.xlsx', import.meta.url),
);

for (const framework of FRAMEWORKS) {
	test(`${framework}: native x14 validation can be edited, saved and cleared`, async ({ page }) => {
		const errors = pageErrors(page);
		await openLanding(page, framework);
		await (await landingFileInput(page)).setInputFiles(fixture);
		await expect(grid(page)).toBeVisible();
		await goToCell(page, 'B2');
		await editor(page)
			.getByRole('button', { name: 'Open the list of values', exact: true })
			.click();
		await expect(
			editor(page).getByRole('listbox', { name: 'List of values' }).getByRole('option'),
		).toHaveText(['red', 'green', 'blue']);
		await editor(page).getByRole('option', { name: 'green', exact: true }).click();
		await ribbon(page).getByRole('tab', { name: 'Data', exact: true }).click();
		const command = ribbon(page).locator('[data-command="data.validation"]').first();
		await clickRibbonControl(page, command);
		const dialog = editor(page).locator('[data-dialog="data-validation"]');
		await expect(dialog.getByLabel('Allow:', { exact: true })).toHaveValue('list');
		await expect(dialog.getByLabel('Source:', { exact: true })).toHaveValue('=Lookup!$A$1:$A$3');
		await dialog.getByLabel('Source:', { exact: true }).fill('=Lookup!$A$2:$A$3');
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(dialog).toBeHidden();
		await editor(page).evaluate(async (node) => {
			const el = node as ValidationEditor;
			await el.load(await el.saveBytes(), 'edited-validation.xlsx');
		});
		await goToCell(page, 'B2');
		await editor(page)
			.getByRole('button', { name: 'Open the list of values', exact: true })
			.click();
		await expect(
			editor(page).getByRole('listbox', { name: 'List of values' }).getByRole('option'),
		).toHaveText(['green', 'blue']);
		await page.keyboard.press('Escape');
		await ribbon(page).getByRole('tab', { name: 'Data', exact: true }).click();
		await clickRibbonControl(page, command);
		await expect(dialog.getByLabel('Source:', { exact: true })).toHaveValue('=Lookup!$A$2:$A$3');
		await dialog.getByRole('button', { name: 'Clear All', exact: true }).click();
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(dialog).toBeHidden();
		await editor(page).evaluate(async (node) => {
			const el = node as ValidationEditor;
			await el.load(await el.saveBytes(), 'cleared-validation.xlsx');
		});
		const ranges = await editor(page).evaluate((node) =>
			(node as ValidationEditor).workbook.sheets
				.find((sheet) => sheet.name === 'Input')!
				.dataValidations.flatMap((rule) => rule.ranges),
		);
		expect(ranges).toEqual([{ start: { row: 2, col: 1 }, end: { row: 4, col: 1 } }]);
		expect(errors).toEqual([]);
	});
}
