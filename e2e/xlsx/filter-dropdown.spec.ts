import { expect, test } from '@playwright/test';
import type { Workbook } from 'ooxml-core/xlsx';
import { editor, FRAMEWORKS, grid, newWorkbook, pageErrors } from './helpers';

interface FilterEditor extends HTMLElement {
	workbook: Workbook;
	select(ref: string): void;
}

for (const framework of FRAMEWORKS) {
	test(`${framework}: the AutoFilter drop-down filters by checklist, Top 10 and custom criteria`, async ({
		page,
	}) => {
		const errors = pageErrors(page);
		await newWorkbook(page, framework);
		await editor(page).evaluate((node) => {
			const element = node as FilterEditor;
			const workbook = element.workbook;
			const sheet = workbook.sheets[0]!;
			[
				['Kind', 'Amount'],
				['apple', 10],
				['banana', 50],
				['apricot', 30],
				['cherry', 20],
			].forEach((values, row) =>
				sheet.rows.set(row, new Map(values.map((value, col) => [col, { value }]))),
			);
			sheet.autoFilter = { range: { start: { row: 0, col: 0 }, end: { row: 4, col: 1 } } };
			element.workbook = workbook;
			element.select('A2');
		});
		const hidden = () =>
			editor(page).evaluate((node) => {
				const sheet = (node as FilterEditor).workbook.sheets[0]!;
				return [1, 2, 3, 4].map((r) => !!sheet.rowInfo.get(r)?.hidden);
			});
		const buttons = grid(page).locator('.xg-fb');
		await expect(buttons).toHaveCount(2);

		// Value checklist: untick "banana".
		await buttons.first().click();
		const menu = grid(page).locator('.xg-fmenu');
		await expect(menu).toBeVisible();
		await menu.locator('.xg-fm-list label', { hasText: 'banana' }).locator('input').uncheck();
		await menu.locator('.xg-fm-ok').click();
		await expect(menu).toHaveCount(0);
		expect(await hidden()).toEqual([false, true, false, false]);
		await expect(buttons.first()).toHaveClass(/xg-fb-active/);

		// Text Filters > Begins With... (Custom AutoFilter dialog).
		await buttons.first().click();
		await menu.getByRole('menuitem', { name: 'Text Filters' }).click();
		await menu.getByRole('menuitem', { name: 'Begins With...' }).click();
		const dialog = editor(page).locator('[data-dialog="custom-autofilter"]');
		await dialog.locator('input[type=text]').fill('ap');
		await dialog.locator('.xve-btn-primary').click();
		expect(await hidden()).toEqual([false, true, false, true]);

		// Number Filters > Top 10 on the Amount column.
		await buttons.nth(1).click();
		await menu.getByRole('menuitem', { name: 'Number Filters' }).click();
		await menu.getByRole('menuitem', { name: 'Top 10...' }).click();
		const top = editor(page).locator('[data-dialog="top10-filter"]');
		await top.locator('input[type=number]').fill('1');
		await top.locator('.xve-btn-primary').click();
		expect(await hidden()).toEqual([true, true, false, true]);

		// Clear Filter From "Amount".
		await buttons.nth(1).click();
		await menu.getByRole('menuitem', { name: /Clear Filter From/ }).click();
		expect(await hidden()).toEqual([false, true, false, true]);
		expect(errors).toEqual([]);
	});
}
