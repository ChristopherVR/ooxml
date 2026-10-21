import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { savedDocument } from './saved-document';
import { clickInEditor, dialogByHeading, fileInput, newDocument } from './helpers';

test('Insert > Shapes draws a floating shape that Shape Format edits and the file keeps', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1700, height: 900 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.locator('.ProseMirror').click();
	await page.keyboard.type('Draw a shape below this line of text.');
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	await editor.getByRole('button', { name: 'Shapes', exact: true }).click();
	const picker = editor.locator('.shape-picker');
	await expect(picker).toBeVisible();
	await expect(picker.locator('h4')).toHaveCount(7);
	await picker.getByRole('button', { name: 'Right Arrow', exact: true }).click();

	// Drag on the page: a rubber band follows the pointer and a shape of that size appears.
	const paragraph = editor.locator('.dve-paper > .ProseMirror > p').first();
	const box = (await paragraph.boundingBox())!;
	await page.mouse.move(box.x + 40, box.y + 8);
	await page.mouse.down();
	await page.mouse.move(box.x + 120, box.y + 40, { steps: 4 });
	await page.mouse.move(box.x + 200, box.y + 70, { steps: 4 });
	await expect(page.locator('.dve-shape-rubber')).toBeVisible();
	await page.mouse.up();
	const shape = editor.locator('.ProseMirror .dve-shape');
	await expect(shape).toHaveCount(1);
	await expect(shape).toHaveClass(/dve-shape-selected/);
	await expect(shape.locator('svg path')).toHaveCount(1);
	const size = (await shape.boundingBox())!;
	expect(size.width).toBeGreaterThan(140);
	expect(size.width).toBeLessThan(180);
	expect(size.height).toBeGreaterThan(50);

	// The contextual Shape Format tab opens with the shape's values.
	const tab = editor.getByRole('tab', { name: 'Shape Format', exact: true });
	await expect(tab).toBeVisible();
	const width = editor.getByRole('spinbutton', { name: 'Shape width (inches)', exact: true });
	await width.fill('3');
	await width.dispatchEvent('change');
	await expect.poll(async () => (await shape.boundingBox())!.width).toBeGreaterThan(280);
	await editor
		.getByRole('combobox', { name: 'Wrap text', exact: true })
		.selectOption({ label: 'Square' });
	await editor.getByRole('button', { name: 'Add text', exact: true }).click();
	const dialog = dialogByHeading(editor, 'Add Text');
	await dialog.getByRole('textbox', { name: 'Text', exact: true }).fill('Hello shape');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(shape.locator('svg text')).toHaveText('Hello shape');

	// Print Layout draws the floating shape on the page.
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	await expect(editor.locator('.dve-print-pages [data-docx-chart="1"] svg').first()).toBeVisible();

	const { xml } = await savedDocument(page, 'ui-shape-drawn.docx');
	expect(xml).toContain('<wp:anchor');
	expect(xml).toContain('prst="rightArrow"');
	expect(xml).toContain('<wp:wrapSquare wrapText="bothSides"/>');
	expect(xml).toContain('<w:t>Hello shape</w:t>');
	expect(xml).toContain('<a:ext cx="2743200"');
});

test('a shape authored in Word is drawn and keeps its markup when text is typed', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'shapes.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await readFile(new URL('./support/shapes.docx', import.meta.url)),
	});
	const editor = page.locator('docx-editor');
	const shapes = editor.locator('.ProseMirror .dve-shape');
	await expect(shapes).toHaveCount(4);
	await expect(editor.locator('.ProseMirror [data-shape-preset="rightArrow"]')).toHaveCount(1);
	await expect(editor.locator('.ProseMirror [data-shape-preset="rect"] svg text')).toHaveText(
		'Red box',
	);
	const first = editor.locator('.dve-paper > .ProseMirror > p').first();
	await clickInEditor(first);
	await page.keyboard.press('End');
	await page.keyboard.type(' typed');
	const { xml } = await savedDocument(page, 'ui-shape-word-edited.docx');
	expect(xml.match(/prst="/g)?.length).toBeGreaterThanOrEqual(4);
	expect(xml).toContain('typed');
});
