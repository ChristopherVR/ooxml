import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import {
	SAMPLE_SHEETS,
	editorProperty,
	formulaBar,
	formulaBarValue,
	goToCell,
	grid,
	landingFileInput,
	nameBox,
	openLanding,
	openSample,
	pageErrors,
	ribbon,
	sheetTabs,
	statusBar,
} from './helpers';

const legacyXls = new URL('./support/legacy-97.xls', import.meta.url);

test('the sample workbook opens with its three sheets and the editor chrome', async ({ page }) => {
	const errors = pageErrors(page);
	await openSample(page);
	await expect(page.locator('#landing')).toBeHidden();
	for (const name of SAMPLE_SHEETS) await expect(sheetTabs(page)).toContainText(name);
	await expect(ribbon(page)).toBeVisible();
	await expect(formulaBar(page)).toBeVisible();
	await expect(nameBox(page)).toBeVisible();
	await expect(statusBar(page)).toBeVisible();
	await expect(grid(page)).toContainText('Regional Sales 2026');
	await expect(grid(page)).toContainText('North');
	expect(await editorProperty<string>(page, 'fileName')).toBe('Sample workbook.xlsx');
	expect(errors).toEqual([]);
});

test('selecting a formula cell shows its formula in the formula bar', async ({ page }) => {
	await openSample(page);
	await goToCell(page, 'N4');
	expect(await formulaBarValue(page)).toBe('=SUM(B4:M4)');
});

test('?sample=1 opens the sample immediately (docs embed)', async ({ page }) => {
	await page.goto('/?sample=1');
	await expect(grid(page)).toBeVisible();
	await expect(sheetTabs(page)).toContainText('Budget');
});

test('a legacy .xls file opens through the landing file input', async ({ page }) => {
	const errors = pageErrors(page);
	await openLanding(page);
	await (
		await landingFileInput(page)
	).setInputFiles({
		name: 'legacy-97.xls',
		mimeType: 'application/vnd.ms-excel',
		buffer: await readFile(legacyXls),
	});
	await expect(grid(page)).toBeVisible();
	for (const name of SAMPLE_SHEETS) await expect(sheetTabs(page)).toContainText(name);
	expect(errors).toEqual([]);
});
