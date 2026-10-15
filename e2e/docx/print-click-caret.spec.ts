import { expect, test, type Page } from '@playwright/test';
import { newDocument, openSample } from './helpers';

test.use({ viewport: { width: 1366, height: 1500 } });

/** A click anywhere on a Print Layout page puts a visible caret at the nearest text, like Word. */
async function enterPrint(page: Page) {
	const editor = page.locator('docx-editor');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('combobox', { name: 'Layout view', exact: true }).selectOption('print');
	await expect(editor.locator('.dve-print-line').first()).toBeVisible();
	return editor;
}

async function clickAndType(page: Page, x: number, y: number, text: string) {
	await page.mouse.click(x, y);
	await expect(page.locator('docx-editor .dve-print-caret')).toHaveCount(1);
	const caret = await page.locator('docx-editor .dve-print-caret').boundingBox();
	expect(caret?.height ?? 0).toBeGreaterThan(4);
	await page.keyboard.type(text);
}

test('blank page: clicks in the margins and empty space place a caret and take typing', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = await enterPrint(page);
	const sheet = (await editor.locator('.dve-print-page').first().boundingBox())!;
	const spots: Array<[string, number, number]> = [
		['bottom blank area', sheet.x + sheet.width / 2, sheet.y + sheet.height * 0.6],
		['right margin', sheet.x + sheet.width - 8, sheet.y + 200],
		['left margin', sheet.x + 8, sheet.y + 300],
		['bottom margin', sheet.x + sheet.width / 2, sheet.y + sheet.height - 10],
		['top margin', sheet.x + sheet.width / 2, sheet.y + 8],
	];
	for (const [label, x, y] of spots) {
		await clickAndType(page, x, y, '.');
		await expect(editor.locator('.ProseMirror').first(), label).toContainText('.');
		await page.keyboard.press('Control+z');
	}
});

test('sample document: clicks right of a line and below the text land on the nearest text', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await openSample(page);
	const editor = await enterPrint(page);
	const line = editor.locator('.dve-print-line').first();
	const box = (await line.boundingBox())!;
	const sheet = (await editor.locator('.dve-print-page').first().boundingBox())!;
	// Right of the end of the first line, inside the page margin.
	await clickAndType(page, sheet.x + sheet.width - 6, box.y + box.height / 2, 'Z');
	await expect(editor.locator('.ProseMirror p').first()).toContainText('Z');
	await page.keyboard.press('Control+z');
	// Immediately after typing (relayout pending), then a quick double click in the bottom margin.
	await page.keyboard.type('Q');
	await page.mouse.click(sheet.x + sheet.width / 2, sheet.y + sheet.height - 12);
	await page.mouse.click(sheet.x + sheet.width / 2, sheet.y + sheet.height - 12);
	await expect(editor.locator('.dve-print-caret')).toHaveCount(1);
});
