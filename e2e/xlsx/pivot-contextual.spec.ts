import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { expect, test, type Page } from '@playwright/test';
import JSZip from 'jszip';
import {
	clickRibbonControl,
	editor,
	formulaBarValue,
	goToCell,
	grid,
	newWorkbook,
	ribbon,
	sheetTabs,
	typeInActiveCell,
} from './helpers';

async function loadPivotWorkbook(page: Page, suffix = '-rows', unstyled = false) {
	let fixture: Uint8Array = readFileSync(
		resolve(process.cwd(), `../../src/core/xlsx/__fixtures__/excel-pivot${suffix}.xlsx`),
	);
	if (unstyled) {
		const zip = await JSZip.loadAsync(fixture);
		const part = 'xl/pivotTables/pivotTable1.xml';
		const xml = await zip.file(part)!.async('string');
		zip.file(part, xml.replace(/<pivotTableStyleInfo[^>]*\/>/, ''));
		fixture = await zip.generateAsync({ type: 'uint8array' });
	}
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

test('shows the requested grand-total row or column', async ({ page }) => {
	await loadPivotWorkbook(page, '');
	await goToCell(page, 'A2');
	const tabs = ribbon(page);
	await tabs.getByRole('tab', { name: 'PivotTable Design' }).click();
	const totals = tabs.getByRole('button', { name: 'Grand Totals', exact: true });
	await clickRibbonControl(page, totals);
	await editor(page).getByRole('menuitemcheckbox', { name: 'On for Rows Only' }).click();
	await goToCell(page, 'D4');
	await expect.poll(() => formulaBarValue(page)).toBe('20');
	await goToCell(page, 'A5');
	await expect.poll(() => formulaBarValue(page)).toBe('');
	await goToCell(page, 'A2');
	await tabs.getByRole('tab', { name: 'PivotTable Design' }).click();
	await clickRibbonControl(page, totals);
	await editor(page).getByRole('menuitemcheckbox', { name: 'On for Columns Only' }).click();
	await goToCell(page, 'C5');
	await expect.poll(() => formulaBarValue(page)).toBe('35');
	await goToCell(page, 'D4');
	await expect.poll(() => formulaBarValue(page)).toBe('');
});

test('leaves an imported report unstyled until a gallery style is selected', async ({ page }) => {
	await loadPivotWorkbook(page, '-rows', true);
	const heading = grid(page)
		.locator('.xg-c:visible')
		.filter({ hasText: /^Category$/ });
	await expect(heading).toHaveCSS('background-color', 'rgba(0, 0, 0, 0)');
	await goToCell(page, 'A2');
	const tabs = ribbon(page);
	await tabs.getByRole('tab', { name: 'PivotTable Design' }).click();
	const gallery = tabs.getByRole('button', { name: 'PivotTable Styles', exact: true });
	await clickRibbonControl(page, gallery);
	await editor(page).getByRole('menuitem', { name: 'Medium Style 4', exact: true }).click();
	await expect(heading).not.toHaveCSS('background-color', 'rgba(0, 0, 0, 0)');
});

test('keeps custom subtotals when moving them to the top and bottom', async ({ page }) => {
	await loadPivotWorkbook(page);
	await goToCell(page, 'A2');
	const tabs = ribbon(page);
	await tabs.getByRole('tab', { name: 'PivotTable Analyze' }).click();
	await clickRibbonControl(page, tabs.locator('[data-command="pivot.field-settings"]'));
	const dialog = editor(page).locator('[data-dialog="pivot-field-settings"]');
	await dialog.locator('input[type=radio][value=custom]').check();
	await dialog.getByLabel('Average', { exact: true }).check();
	await dialog.getByLabel('Count', { exact: true }).check();
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await tabs.getByRole('tab', { name: 'PivotTable Design' }).click();
	for (const position of ['Top', 'Bottom']) {
		await clickRibbonControl(page, tabs.getByRole('button', { name: 'Subtotals', exact: true }));
		await editor(page)
			.getByRole('menuitem', { name: `Show All Subtotals at ${position}` })
			.click();
		await expect(
			grid(page)
				.locator('.xg-c:visible')
				.filter({ hasText: /^Fruit Average$/ }),
		).toBeVisible();
		await expect(
			grid(page)
				.locator('.xg-c:visible')
				.filter({ hasText: /^Fruit Count$/ }),
		).toBeVisible();
	}
});
