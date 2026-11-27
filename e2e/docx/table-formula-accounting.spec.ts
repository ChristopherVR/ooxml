import { expect, test } from '@playwright/test';
import { clickInEditor, dialogByHeading, insertTableOfSize, newDocument, reveal } from './helpers';

test('Formula sums the Word numeric range above a cell with accounting text nearby', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1700, height: 960 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	await insertTableOfSize(page, 3, 1);
	const editor = page.locator('docx-editor');
	const cells = editor.locator('.ProseMirror td');
	for (const [index, value] of ['(5)', '9'].entries()) {
		await clickInEditor(cells.nth(index));
		await page.keyboard.type(value);
	}
	await clickInEditor(cells.nth(2));
	await editor.getByRole('tab', { name: 'Table Layout', exact: true }).click();
	const command = editor.getByRole('button', { name: 'Formula', exact: true });
	await reveal(editor, command);
	await command.click();
	const formula = dialogByHeading(editor, 'Formula');
	await expect(formula.getByRole('textbox', { name: 'Formula', exact: true })).toHaveValue(
		'=SUM(ABOVE)',
	);
	await formula.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(cells.nth(2)).toHaveText('9');
});
