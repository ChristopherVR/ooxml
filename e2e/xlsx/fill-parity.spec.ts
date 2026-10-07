import { expect, test } from '@playwright/test';
import { editor, grid, newWorkbook, typeInActiveCell } from './helpers';

test('Fill Down and Fill Right copy weekday text rather than generating a series', async ({
	page,
}) => {
	await newWorkbook(page);
	await typeInActiveCell(page, 'Monday');
	await editor(page).evaluate((node) =>
		(node as unknown as { select(ref: string): void }).select('A1:A3'),
	);
	await grid(page).focus();
	await page.keyboard.press('Control+D');
	const values = () =>
		editor(page).evaluate((node) => {
			const sheet = (
				node as unknown as {
					workbook: { sheets: { rows: Map<number, Map<number, { value: unknown }>> }[] };
				}
			).workbook.sheets[0]!;
			return [
				sheet.rows.get(1)?.get(0)?.value ?? null,
				sheet.rows.get(2)?.get(0)?.value ?? null,
				sheet.rows.get(0)?.get(1)?.value ?? null,
				sheet.rows.get(0)?.get(2)?.value ?? null,
			];
		});
	await expect.poll(values).toEqual(['Monday', 'Monday', null, null]);
	await page.keyboard.press('Control+Z');
	await expect.poll(values).toEqual([null, null, null, null]);
	await page.keyboard.press('Control+Y');
	await expect.poll(values).toEqual(['Monday', 'Monday', null, null]);
	await editor(page).evaluate((node) =>
		(node as unknown as { select(ref: string): void }).select('A1:C1'),
	);
	await grid(page).focus();
	await page.keyboard.press('Control+R');
	await expect.poll(values).toEqual(['Monday', 'Monday', 'Monday', 'Monday']);
});
