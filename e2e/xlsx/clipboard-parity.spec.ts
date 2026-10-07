import { expect, test } from '@playwright/test';
import { editor, goToCell, grid, newWorkbook, ribbon, typeInActiveCell } from './helpers';

test('native clipboard copies a hyperlink target and undo restores the destination link', async ({
	page,
}) => {
	await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
	await newWorkbook(page);
	for (const [ref, text, url] of [
		['A1', '2', 'source'],
		['C1', '9', 'destination'],
	]) {
		await goToCell(page, ref!);
		await typeInActiveCell(page, text!);
		await goToCell(page, ref!);
		await page.keyboard.press('Control+K');
		const dialog = editor(page).locator('[data-dialog="hyperlink"]');
		await dialog.getByLabel('Address:', { exact: true }).fill(`https://example.com/${url}`);
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	}
	await goToCell(page, 'A1');
	await page.keyboard.press('Control+C');
	await goToCell(page, 'C1');
	await page.keyboard.press('Control+V');
	const read = () =>
		editor(page).evaluate((node) => {
			const sheet = (
				node as unknown as {
					workbook: {
						sheets: {
							hyperlinks: {
								range: { start: { row: number; col: number }; end: { row: number; col: number } };
								target?: string;
							}[];
							rows: Map<number, Map<number, { value: unknown }>>;
						}[];
					};
				}
			).workbook.sheets[0]!;
			return {
				value: sheet.rows.get(0)?.get(2)?.value,
				target: sheet.hyperlinks.find(
					(link) => link.range.start.row === 0 && link.range.start.col === 2,
				)?.target,
			};
		});
	await expect.poll(read).toEqual({ value: 2, target: 'https://example.com/source' });
	await page.keyboard.press('Control+Z');
	await expect.poll(read).toEqual({ value: 9, target: 'https://example.com/destination' });
	await page.keyboard.press('Control+Y');
	await expect.poll(read).toEqual({ value: 2, target: 'https://example.com/source' });
});

test('Validation paste copies a blank cell rule and enforces it without replacing the value', async ({
	page,
}) => {
	await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
	await newWorkbook(page);
	await ribbon(page).getByRole('tab', { name: 'Data', exact: true }).click();
	await ribbon(page).locator('[data-command="data.validation"]').first().click();
	const settings = editor(page).locator('[data-dialog="data-validation"]');
	await settings.getByLabel('Allow:', { exact: true }).selectOption('whole');
	await settings.getByLabel('Minimum:', { exact: true }).fill('1');
	await settings.getByLabel('Maximum:', { exact: true }).fill('5');
	await settings.getByRole('button', { name: 'OK', exact: true }).click();
	await goToCell(page, 'C1');
	await typeInActiveCell(page, '9');
	await goToCell(page, 'A1');
	await page.keyboard.press('Control+C');
	await goToCell(page, 'C1');
	await page.keyboard.press('Control+Alt+V');
	const dialog = editor(page).locator('[data-dialog="paste-special"]');
	await dialog.getByLabel('Validation', { exact: true }).check();
	await dialog.getByLabel('Skip blanks', { exact: true }).check();
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	const value = () =>
		editor(page).evaluate(
			(node) =>
				(
					node as unknown as {
						workbook: { sheets: { rows: Map<number, Map<number, { value: unknown }>> }[] };
					}
				).workbook.sheets[0]!.rows.get(0)?.get(2)?.value,
		);
	await expect.poll(value).toBe(9);
	await typeInActiveCell(page, '10');
	await expect.poll(value).toBe(9);
	await page.keyboard.press('Escape');
	await goToCell(page, 'C1');
	await typeInActiveCell(page, '4');
	await expect.poll(value).toBe(4);
});

test('Comments paste copies a blank cell note through the native clipboard and undoes once', async ({
	page,
}) => {
	await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
	await newWorkbook(page);
	await page.keyboard.press('Shift+F2');
	const commentDialog = editor(page).locator('[data-dialog="comment"]');
	await commentDialog.getByLabel('Comment', { exact: true }).fill('source note');
	await commentDialog.getByRole('button', { name: 'OK', exact: true }).click();
	await goToCell(page, 'C1');
	await typeInActiveCell(page, '9');
	await goToCell(page, 'A1');
	await page.keyboard.press('Control+C');
	await goToCell(page, 'C1');
	await page.keyboard.press('Control+Alt+V');
	const dialog = editor(page).locator('[data-dialog="paste-special"]');
	await dialog.getByLabel('Comments', { exact: true }).check();
	await dialog.getByLabel('Skip blanks', { exact: true }).check();
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	const read = () =>
		editor(page).evaluate((node) => {
			const sheet = (
				node as unknown as {
					workbook: {
						sheets: {
							comments: { address: { row: number; col: number }; text: string }[];
							rows: Map<number, Map<number, { value: unknown }>>;
						}[];
					};
				}
			).workbook.sheets[0]!;
			return {
				note: sheet.comments.find((c) => c.address.row === 0 && c.address.col === 2)?.text ?? null,
				value: sheet.rows.get(0)?.get(2)?.value,
			};
		});
	await expect.poll(read).toEqual({ note: 'source note', value: 9 });
	await page.keyboard.press('Control+Z');
	await expect.poll(read).toEqual({ note: null, value: 9 });
	await page.keyboard.press('Control+Y');
	await expect.poll(read).toEqual({ note: 'source note', value: 9 });
});

test('Column Widths uses the native clipboard, retains content and undoes once', async ({
	page,
}) => {
	await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
	await newWorkbook(page);
	await goToCell(page, 'A1');
	await ribbon(page).getByRole('button', { name: 'Format', exact: true }).click();
	await editor(page).getByRole('menuitem', { name: 'Column Width...', exact: true }).click();
	const widthDialog = editor(page).locator('[data-dialog="column-width"]');
	await widthDialog.getByLabel('Column width:', { exact: true }).fill('20');
	await widthDialog.getByRole('button', { name: 'OK', exact: true }).click();
	await goToCell(page, 'C1');
	await typeInActiveCell(page, '9');
	const read = () =>
		editor(page).evaluate((node) => {
			const sheet = (
				node as unknown as {
					workbook: {
						sheets: {
							columns: { min: number; max: number; width?: number }[];
							rows: Map<number, Map<number, { value: unknown }>>;
						}[];
					};
				}
			).workbook.sheets[0]!;
			return {
				width: sheet.columns.find((c) => c.min <= 2 && c.max >= 2)?.width ?? null,
				value: sheet.rows.get(0)?.get(2)?.value,
			};
		});
	const before = await read();
	await goToCell(page, 'A1');
	await page.keyboard.press('Control+C');
	await goToCell(page, 'C1');
	await page.keyboard.press('Control+Alt+V');
	const dialog = editor(page).locator('[data-dialog="paste-special"]');
	await dialog.getByLabel('Column widths', { exact: true }).check();
	await dialog.getByLabel('Skip blanks', { exact: true }).check();
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect.poll(read).toEqual({ width: 20, value: 9 });
	await page.keyboard.press('Control+Z');
	await expect.poll(read).toEqual(before);
	await page.keyboard.press('Control+Y');
	await expect.poll(read).toEqual({ width: 20, value: 9 });
});

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

test('All Except Borders copies font and value while retaining the destination border', async ({
	page,
}) => {
	await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
	await newWorkbook(page);
	await typeInActiveCell(page, '2');
	await goToCell(page, 'A1');
	await ribbon(page).locator('[data-command="home.bold"]').first().click();
	await goToCell(page, 'C1');
	await typeInActiveCell(page, '9');
	await goToCell(page, 'C1');
	await ribbon(page).locator('[data-command="home.borders"]').first().click();
	const result = () =>
		editor(page).evaluate((node) => {
			const book = (
				node as unknown as {
					workbook: {
						styles: { font: { bold?: boolean }; border: { bottom?: { style: string } } }[];
						sheets: { rows: Map<number, Map<number, { value: unknown; styleId?: number }>> }[];
					};
				}
			).workbook;
			const cell = book.sheets[0]!.rows.get(0)?.get(2);
			const style = book.styles[cell?.styleId ?? 0]!;
			return {
				value: cell?.value,
				bold: style.font.bold ?? false,
				border: style.border.bottom?.style,
			};
		});
	await expect.poll(result).toEqual({ value: 9, bold: false, border: 'thin' });
	await goToCell(page, 'A1');
	await page.keyboard.press('Control+C');
	await goToCell(page, 'C1');
	await page.keyboard.press('Control+Alt+V');
	const dialog = editor(page).locator('[data-dialog="paste-special"]');
	await dialog.getByLabel('All except borders', { exact: true }).check();
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect.poll(result).toEqual({ value: 2, bold: true, border: 'thin' });
	await page.keyboard.press('Control+Z');
	await expect.poll(result).toEqual({ value: 9, bold: false, border: 'thin' });
});

test('Paste Special Multiply applies native clipboard values and undoes once', async ({ page }) => {
	await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
	await newWorkbook(page);
	await typeInActiveCell(page, '2');
	await goToCell(page, 'C1');
	await typeInActiveCell(page, '=5+5');
	await goToCell(page, 'A1');
	await page.keyboard.press('Control+C');
	await goToCell(page, 'C1');
	await page.keyboard.press('Control+Alt+V');
	const dialog = editor(page).locator('[data-dialog="paste-special"]');
	await dialog.getByLabel('Values', { exact: true }).check();
	await dialog.getByLabel('Multiply', { exact: true }).check();
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	const result = () =>
		editor(page).evaluate((node) => {
			const sheet = (
				node as unknown as {
					workbook: {
						sheets: { rows: Map<number, Map<number, { value: unknown; formula?: string }>> }[];
					};
				}
			).workbook.sheets[0]!;
			const cell = sheet.rows.get(0)?.get(2);
			return [cell?.value, cell?.formula];
		});
	await expect.poll(result).toEqual([20, '(5+5)*2']);
	await page.keyboard.press('Control+Z');
	await expect.poll(result).toEqual([10, '5+5']);
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
