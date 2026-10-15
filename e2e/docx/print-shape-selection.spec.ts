import { expect, test, type Locator, type Page } from '@playwright/test';
import { savedDocument } from './saved-document';
import { dialogByHeading, newDocument } from './helpers';

test.use({ viewport: { width: 1700, height: 1000 } });

/** A new document with one floating rectangle drawn in Draft view, then Print Layout. */
async function openWithShape(page: Page): Promise<Locator> {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.locator('.ProseMirror').click();
	await page.keyboard.type('Draw a shape below this line of text.');
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	await editor.getByRole('button', { name: 'Shapes', exact: true }).click();
	await editor
		.locator('.shape-picker')
		.getByRole('button', { name: 'Rectangle', exact: true })
		.click();
	const paragraph = editor.locator('.dve-paper > .ProseMirror > p').first();
	const box = (await paragraph.boundingBox())!;
	await page.mouse.move(box.x + 40, box.y + 8);
	await page.mouse.down();
	await page.mouse.move(box.x + 120, box.y + 40, { steps: 4 });
	await page.mouse.move(box.x + 200, box.y + 70, { steps: 4 });
	await page.mouse.up();
	await expect(editor.locator('.ProseMirror .dve-shape')).toHaveCount(1);
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	await expect(editor.locator('.dve-print-float')).toHaveCount(1);
	return editor;
}

const float = (editor: Locator) => editor.locator('.dve-print-float[data-block-id]');
const frame = (editor: Locator) => editor.locator('.dve-print-float-frame');

async function rect(locator: Locator) {
	let box = await locator.boundingBox();
	for (let attempt = 0; !box && attempt < 20; attempt++) {
		await locator.page().waitForTimeout(25);
		box = await locator.boundingBox();
	}
	return box!;
}

test('a floating shape is selected on the page with handles and the Shape Format tab', async ({
	page,
}) => {
	const editor = await openWithShape(page);
	const box = await rect(float(editor));
	await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
	await expect(frame(editor)).toHaveCount(1);
	await expect(frame(editor).locator('[data-float-handle]')).toHaveCount(8);
	await expect(editor.getByRole('tab', { name: 'Shape Format', exact: true })).toBeVisible();
});

test('dragging and resizing a shape on the page change its saved geometry', async ({ page }) => {
	const editor = await openWithShape(page);
	const before = await rect(float(editor));
	const startX = before.x + before.width / 2;
	const startY = before.y + before.height / 2;
	await page.mouse.move(startX, startY);
	await page.mouse.down();
	await page.mouse.move(startX + 20, startY + 10, { steps: 4 });
	await page.mouse.move(startX + 50, startY + 30, { steps: 4 });
	await page.mouse.up();
	await expect.poll(async () => Math.round((await rect(float(editor))).x - before.x)).toBe(50);
	expect(Math.round((await rect(float(editor))).y - before.y)).toBe(30);
	// An edge handle changes one dimension only (no aspect lock on shapes).
	const handle = await rect(frame(editor).locator('[data-float-handle="e"]'));
	const hx = handle.x + handle.width / 2;
	const hy = handle.y + handle.height / 2;
	await page.mouse.move(hx, hy);
	await page.mouse.down();
	await page.mouse.move(hx + 30, hy, { steps: 4 });
	await page.mouse.move(hx + 60, hy, { steps: 4 });
	await page.mouse.up();
	await expect
		.poll(async () => Math.round((await rect(float(editor))).width - before.width))
		.toBe(60);
	expect(Math.round((await rect(float(editor))).height - before.height)).toBe(0);
	const { xml } = await savedDocument(page, 'ui-print-shape-moved.docx');
	expect(xml).toContain('prst="rect"');
	expect(xml).toContain('<wp:anchor');
});

test('Delete removes the selected shape in one undo step', async ({ page }) => {
	const editor = await openWithShape(page);
	const box = await rect(float(editor));
	await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
	await expect(frame(editor)).toHaveCount(1);
	// The draw closes history only before the next change; let the history's group delay pass.
	await page.waitForTimeout(600);
	await page.keyboard.press('Delete');
	await expect(editor.locator('.dve-print-float')).toHaveCount(0);
	await page.keyboard.press('Control+z');
	await expect(editor.locator('.dve-print-float')).toHaveCount(1);
});

test('double-clicking a shape edits its text on the page and the file keeps it', async ({
	page,
}) => {
	const editor = await openWithShape(page);
	await editor.getByRole('tab', { name: 'Shape Format', exact: true }).click();
	const box = await rect(float(editor));
	await page.mouse.dblclick(box.x + box.width / 2, box.y + box.height / 2);
	const field = editor.locator('.dve-print-float-text');
	await expect(field).toBeVisible();
	await field.fill('Typed on the page');
	await field.press('Control+Enter');
	await expect(field).toHaveCount(0);
	await expect(float(editor).locator('svg text')).toHaveText('Typed on the page');
	const { xml } = await savedDocument(page, 'ui-print-shape-text.docx');
	expect(xml).toContain('<w:t>Typed on the page</w:t>');
	// A press inside a selected shape that holds text puts the caret back in it.
	const again = await rect(float(editor));
	await page.mouse.click(again.x + again.width / 2, again.y + again.height / 2);
	await expect(field).toBeVisible();
	await field.press('Escape');
	await expect(field).toHaveCount(0);
	// The Shape Format dialog sees the same text.
	await editor.getByRole('button', { name: 'Add text', exact: true }).click();
	const dialog = dialogByHeading(editor, 'Add Text');
	await expect(dialog.getByRole('textbox', { name: 'Text', exact: true })).toHaveValue(
		'Typed on the page',
	);
});
