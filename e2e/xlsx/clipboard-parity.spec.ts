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

test('Paste Special combines values, transpose and Skip Blanks with one undo', async ({ page }) => {
	await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
	await newWorkbook(page);
	const select = (ref: string) =>
		editor(page).evaluate(
			(node, value) => (node as unknown as { select(ref: string): void }).select(value),
			ref,
		);
	await select('A2');
	await typeInActiveCell(page, '=""');
	await select('A3');
	await typeInActiveCell(page, '0');
	for (const ref of ['C1', 'D1', 'E1']) {
		await select(ref);
		await typeInActiveCell(page, '9');
	}
	await select('A1:A3');
	await grid(page).focus();
	await page.keyboard.press('Control+C');
	await select('C1');
	await grid(page).focus();
	await page.keyboard.press('Control+Alt+V');
	const dialog = editor(page).locator('[data-dialog="paste-special"]');
	await expect(dialog).toBeVisible();
	await dialog.getByLabel('Values', { exact: true }).check();
	await dialog.getByLabel('Skip blanks', { exact: true }).check();
	await dialog.getByLabel('Transpose', { exact: true }).check();
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	const values = () =>
		editor(page).evaluate((node) => {
			const sheet = (
				node as unknown as {
					workbook: {
						sheets: { rows: Map<number, Map<number, { value: unknown; formula?: string }>> }[];
					};
				}
			).workbook.sheets[0]!;
			return [2, 3, 4].map((col) => ({
				value: sheet.rows.get(0)?.get(col)?.value ?? null,
				formula: sheet.rows.get(0)?.get(col)?.formula ?? null,
			}));
		});
	await expect.poll(values).toEqual([
		{ value: 9, formula: null },
		{ value: '', formula: null },
		{ value: 0, formula: null },
	]);
	await grid(page).focus();
	await page.keyboard.press('Control+Z');
	await expect.poll(values).toEqual([2, 3, 4].map(() => ({ value: 9, formula: null })));
});
