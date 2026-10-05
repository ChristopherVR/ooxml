import { expect, test } from '@playwright/test';
import {
	editor,
	editorProperty,
	formulaBarValue,
	goToCell,
	grid,
	nameBoxValue,
	newWorkbook,
	pageErrors,
	sheetTabs,
	typeInActiveCell,
} from './helpers';

test('a new workbook starts as Book1.xlsx with one sheet and A1 selected', async ({ page }) => {
	const errors = pageErrors(page);
	await newWorkbook(page);
	await expect(sheetTabs(page)).toContainText('Sheet1');
	expect(await nameBoxValue(page)).toBe('A1');
	expect(await editorProperty<string>(page, 'fileName')).toBe('Book1.xlsx');
	expect(errors).toEqual([]);
});

test('typing into the grid enters values and formulas that recalculate', async ({ page }) => {
	await newWorkbook(page);
	await typeInActiveCell(page, '12');
	await typeInActiveCell(page, '30');
	await typeInActiveCell(page, '=SUM(A1:A2)');
	await expect(grid(page)).toContainText('42');
	await goToCell(page, 'A3');
	expect(await formulaBarValue(page)).toBe('=SUM(A1:A2)');
	await goToCell(page, 'A1');
	await typeInActiveCell(page, '20');
	await expect(grid(page)).toContainText('50');
	expect(await editorProperty<boolean>(page, 'dirty')).toBe(true);
});

test('save produces an .xlsx workbook', async ({ page }) => {
	await newWorkbook(page);
	await typeInActiveCell(page, 'Saved value');
	const bytes = await editor(page).evaluate(async (node) => {
		const blob = await (node as unknown as { save(): Promise<Blob> }).save();
		return Array.from(new Uint8Array(await blob.arrayBuffer()).slice(0, 4));
	});
	expect(bytes).toEqual([0x50, 0x4b, 0x03, 0x04]);
});
