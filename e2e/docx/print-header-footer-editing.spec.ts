import { expect, test, type Locator, type Page } from '@playwright/test';
import { fileInput, newDocument } from './helpers';
import { multiPageDocx } from './print-click-fixture';

test.use({ viewport: { width: 1366, height: 1500 } });

async function open(page: Page): Promise<Locator> {
	page.on('dialog', (dialog) => void dialog.accept());
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'multi.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await multiPageDocx(),
	});
	const editor = page.locator('docx-editor');
	await expect(editor.locator('.ProseMirror').first()).toContainText('Page one intro');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	await expect(editor.locator('.dve-print-line').first()).toBeVisible();
	return editor;
}

test('double-clicking a page header edits it in place and every page shows the edit', async ({
	page,
}) => {
	const editor = await open(page);
	const header = editor.locator('.dve-print-page').first().locator('.dve-print-header');
	await expect(header).toContainText('Header text');
	await header.dblclick({ force: true });
	const story = header.locator('.dve-header-footer-editor .ProseMirror');
	await expect(story).toBeFocused();
	// The body dims and the Header & Footer tab offers the exit.
	await expect(editor.locator('.dve-print-pages').first()).toHaveClass(/dve-print-story-editing/);
	await expect(editor.getByRole('button', { name: 'Close Header and Footer' })).toBeVisible();
	await page.keyboard.press('End');
	await page.keyboard.type(' edited');
	await page.keyboard.press('Escape');
	await expect(editor.locator('.dve-print-pages').first()).not.toHaveClass(
		/dve-print-story-editing/,
	);
	for (const sheet of [0, 1, 2])
		await expect(
			editor.locator('.dve-print-page').nth(sheet).locator('.dve-print-header'),
		).toContainText('Header text edited');
});

test('footer: the Close button exits; a double-click on the body exits too', async ({ page }) => {
	const editor = await open(page);
	const footer = editor.locator('.dve-print-page').first().locator('.dve-print-footer');
	await footer.dblclick({ force: true });
	await expect(footer.locator('.ProseMirror')).toBeFocused();
	await page.keyboard.press('End');
	await page.keyboard.type('!');
	await editor.getByRole('button', { name: 'Close Header and Footer' }).click();
	await expect(editor.locator('.dve-print-pages').first()).not.toHaveClass(
		/dve-print-story-editing/,
	);
	await expect(editor.locator('.dve-print-page').nth(1).locator('.dve-print-footer')).toContainText(
		'Footer text!',
	);
	// A single click on the body while editing does not move the caret into it.
	await footer.dblclick({ force: true });
	await expect(footer.locator('.ProseMirror')).toBeFocused();
	const body = (await editor
		.locator('.dve-print-line', { hasText: 'Page one intro' })
		.boundingBox())!;
	await page.mouse.click(body.x + 4, body.y + body.height / 2);
	await expect(footer.locator('.ProseMirror')).toBeFocused();
	await page.mouse.dblclick(body.x + 4, body.y + body.height / 2);
	await expect(editor.locator('.dve-print-pages').first()).not.toHaveClass(
		/dve-print-story-editing/,
	);
	await page.keyboard.type('Q');
	await expect(editor.locator('.ProseMirror').first()).toContainText('Q');
});

test('a blank document gets a header by double-clicking the top margin', async ({ page }) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	const sheet = (await editor.locator('.dve-print-page').first().boundingBox())!;
	await page.mouse.dblclick(sheet.x + sheet.width / 2, sheet.y + 40);
	await expect(editor.locator('.dve-print-header .ProseMirror')).toBeFocused();
	await page.keyboard.type('New header');
	await page.keyboard.press('Escape');
	await expect(
		editor.locator('.dve-print-page').first().locator('.dve-print-header'),
	).toContainText('New header');
});
