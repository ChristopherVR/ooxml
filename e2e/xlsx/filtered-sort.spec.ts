import { expect, test } from '@playwright/test';
import { getCell, loadXlsx, type Workbook } from 'ooxml-core/xlsx';
import { editor, FRAMEWORKS, newWorkbook, pageErrors, ribbon } from './helpers';

interface SortEditor extends HTMLElement {
	workbook: Workbook;
	select(ref: string): void;
	undo(): void;
	redo(): void;
	saveBytes(): Promise<Uint8Array>;
}

for (const framework of FRAMEWORKS) {
	test(`${framework}: ribbon sorting preserves filtered records and subtotals`, async ({
		page,
	}) => {
		const errors = pageErrors(page);
		await newWorkbook(page, framework);
		await editor(page).evaluate((node) => {
			const element = node as SortEditor;
			const workbook = element.workbook;
			const sheet = workbook.sheets[0]!;
			for (const [row, values] of [
				['Kind', 'Amount'],
				['keep', 10],
				['drop', 50],
				['keep', 30],
			].entries()) {
				sheet.rows.set(row, new Map(values.map((value, col) => [col, { value }])));
			}
			sheet.rows.set(5, new Map([[1, { value: 40, formula: 'SUBTOTAL(9,B2:B4)' }]]));
			sheet.rowInfo.set(2, { hidden: true, filteredOut: true });
			sheet.autoFilter = {
				range: { start: { row: 0, col: 0 }, end: { row: 3, col: 1 } },
				columns: [{ offset: 0, values: ['keep'] }],
			};
			element.workbook = workbook;
			element.select('B2');
		});
		await ribbon(page).getByRole('tab', { name: 'Data', exact: true }).click();
		await ribbon(page).locator('[data-command="data.sort-desc"]').first().click();
		const observed = await editor(page).evaluate((node) => {
			const sheet = (node as SortEditor).workbook.sheets[0]!;
			return {
				amounts: [1, 2, 3].map((r) => sheet.rows.get(r)?.get(1)?.value),
				hidden: [1, 2, 3].map((r) => !!sheet.rowInfo.get(r)?.hidden),
				subtotal: sheet.rows.get(5)?.get(1)?.value,
			};
		});
		expect(observed).toEqual({ amounts: [30, 50, 10], hidden: [false, true, false], subtotal: 40 });
		await editor(page).evaluate((node) => (node as SortEditor).undo());
		expect(
			await editor(page).evaluate(
				(node) => (node as SortEditor).workbook.sheets[0]!.rows.get(1)?.get(1)?.value,
			),
		).toBe(10);
		await editor(page).evaluate((node) => (node as SortEditor).redo());
		const bytes = await editor(page).evaluate(async (node) =>
			Array.from(await (node as SortEditor).saveBytes()),
		);
		const loaded = await loadXlsx(new Uint8Array(bytes));
		expect([1, 2, 3].map((r) => getCell(loaded.sheets[0]!, r, 1)?.value)).toEqual([30, 50, 10]);
		expect(loaded.sheets[0]!.rowInfo.get(2)?.hidden).toBe(true);
		expect(getCell(loaded.sheets[0]!, 5, 1)?.value).toBe(40);
		expect(errors).toEqual([]);
	});
}
