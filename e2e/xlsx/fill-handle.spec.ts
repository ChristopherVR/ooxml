import { expect, test } from '@playwright/test';
import type { Workbook } from 'ooxml-core/xlsx';
import { editor, FRAMEWORKS, grid, newWorkbook, typeInActiveCell } from './helpers';

interface FillEditor extends HTMLElement {
	workbook: Workbook;
	select(ref: string): void;
}

for (const framework of FRAMEWORKS) {
	test(`${framework}: dragging the fill handle extends a numeric series in one undo step`, async ({
		page,
	}) => {
		await newWorkbook(page, framework);
		await typeInActiveCell(page, '1');
		await typeInActiveCell(page, '2');
		await editor(page).evaluate((node) => (node as FillEditor).select('A1:A2'));
		const handle = editor(page).locator('.xg-q-main [data-handle="fill"]');
		await expect(handle).toBeVisible();
		const box = (await handle.boundingBox())!;
		const rowHeight = await editor(page).evaluate((node) => {
			const sheet = (node as FillEditor).workbook.sheets[0]!;
			return sheet.defaultRowHeight * (96 / 72) * (sheet.view.zoom / 100);
		});
		const x = box.x + box.width / 2;
		const y = box.y + box.height / 2;
		await page.mouse.move(x, y);
		await page.mouse.down();
		await page.mouse.move(x, y + 3 * rowHeight - 3, { steps: 8 });
		await page.mouse.up();
		const values = () =>
			editor(page).evaluate((node) => {
				const rows = (node as FillEditor).workbook.sheets[0]!.rows;
				return [0, 1, 2, 3, 4].map((row) => rows.get(row)?.get(0)?.value ?? null);
			});
		await expect.poll(values).toEqual([1, 2, 3, 4, 5]);
		await grid(page).focus();
		await page.keyboard.press('Control+Z');
		await expect.poll(values).toEqual([1, 2, null, null, null]);
		await page.keyboard.press('Control+Y');
		await expect.poll(values).toEqual([1, 2, 3, 4, 5]);
	});
}
