import { expect, test } from '@playwright/test';
import type { Workbook } from 'ooxml-core/xlsx';
import {
	editor,
	FRAMEWORKS,
	goToCell,
	newWorkbook,
	openLanding,
	pageErrors,
	typeInActiveCell,
} from './helpers';

interface FormulaEditor extends HTMLElement {
	workbook: Workbook;
	saveBytes(): Promise<Uint8Array>;
	load(bytes: Uint8Array, name: string): Promise<void>;
}

for (const framework of FRAMEWORKS) {
	test(`${framework}: FILTERXML spills, recalculates and survives save`, async ({ page }) => {
		const errors = pageErrors(page);
		await openLanding(page, framework);
		await newWorkbook(page, framework);
		await goToCell(page, 'A1');
		await typeInActiveCell(page, '<r><a>001</a><a>TRUE</a><a>tail</a></r>');
		await goToCell(page, 'C1');
		await typeInActiveCell(page, '=FILTERXML(A1,"//a")');
		const values = () =>
			editor(page).evaluate((node) => {
				const sheet = (node as FormulaEditor).workbook.sheets[0]!;
				return [0, 1, 2].map((row) => sheet.rows.get(row)?.get(2)?.value ?? null);
			});
		await expect.poll(values).toEqual([1, true, 'tail']);
		await goToCell(page, 'A1');
		await typeInActiveCell(page, '<r><a>changed</a></r>');
		await expect.poll(values).toEqual(['changed', null, null]);
		await page.keyboard.press('Control+z');
		await expect.poll(values).toEqual([1, true, 'tail']);
		await editor(page).evaluate(async (node) => {
			const el = node as FormulaEditor;
			await el.load(await el.saveBytes(), 'filterxml.xlsx');
		});
		await expect.poll(values).toEqual([1, true, 'tail']);
		expect(errors).toEqual([]);
	});
}
