import { expect, test } from '@playwright/test';
import { editor, grid, newWorkbook, typeInActiveCell } from './helpers';

test('native clipboard repeats one copied cell across the selected rectangle', async ({ page }) => {
	await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
	await newWorkbook(page);
	await typeInActiveCell(page, 'Monday');
	const select = (ref: string) =>
		editor(page).evaluate(
			(node, value) => (node as unknown as { select(ref: string): void }).select(value),
			ref,
		);
	await select('A1');
	await grid(page).focus();
	await page.keyboard.press('Control+C');
	await select('C3:D4');
	await grid(page).focus();
	await page.keyboard.press('Control+V');
	await expect
		.poll(() =>
			editor(page).evaluate((node) => {
				const sheet = (
					node as unknown as {
						workbook: { sheets: { rows: Map<number, Map<number, { value: unknown }>> }[] };
					}
				).workbook.sheets[0]!;
				return [2, 3].flatMap((row) =>
					[2, 3].map((col) => sheet.rows.get(row)?.get(col)?.value ?? null),
				);
			}),
		)
		.toEqual(['Monday', 'Monday', 'Monday', 'Monday']);
});
