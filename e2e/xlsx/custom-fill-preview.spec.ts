import { expect, test } from '@playwright/test';
import { FRAMEWORKS, editor, goToCell, newWorkbook, typeInActiveCell } from './helpers';

for (const framework of FRAMEWORKS) {
	test(`${framework}: custom fill previews stay in the dialog and submit as one undo step`, async ({
		page,
	}) => {
		await newWorkbook(page, framework);
		await typeInActiveCell(page, 'preview');
		await goToCell(page, 'A1');
		const fill = () =>
			editor(page).evaluate((node) => {
				const book = (
					node as unknown as {
						workbook: {
							styles: { fill: unknown }[];
							sheets: { rows: Map<number, Map<number, { styleId?: number }>> }[];
						};
					}
				).workbook;
				return book.styles[book.sheets[0]!.rows.get(0)?.get(0)?.styleId ?? 0]!.fill;
			});
		const before = await fill();
		await page.keyboard.press('Control+1');
		let dialog = editor(page).locator('[data-dialog="format-cells"]');
		await dialog.getByRole('tab', { name: 'Fill', exact: true }).click();
		const input = dialog.locator(
			'office-ui-color-grid[aria-label="Background Color:"] input[type="color"]',
		);
		await input.evaluate((node) => {
			const field = node as HTMLInputElement;
			for (const color of ['#123456', '#abcdef']) {
				field.value = color;
				field.dispatchEvent(new Event('input'));
			}
		});
		await expect(dialog.locator('[data-fill-preview]')).toHaveCSS(
			'background-color',
			'rgb(171, 205, 239)',
		);
		expect(await fill()).toEqual(before);
		await input.dispatchEvent('change');
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		expect(await fill()).toMatchObject({ fgColor: { rgb: 'FFABCDEF' } });
		await page.keyboard.press('Control+z');
		expect(await fill()).toEqual(before);
		await page.keyboard.press('Control+1');
		dialog = editor(page).locator('[data-dialog="format-cells"]');
		await dialog.getByRole('tab', { name: 'Fill', exact: true }).click();
		await dialog
			.locator('input[type="color"]')
			.first()
			.evaluate((node) => {
				(node as HTMLInputElement).value = '#112233';
				node.dispatchEvent(new Event('input'));
			});
		await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
		expect(await fill()).toEqual(before);
	});
}
