import { expect, test, type Locator, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';

test.use({ viewport: { width: 1366, height: 1500 } });

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function fixture(): Promise<Buffer> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body>
<w:p><w:r><w:t>Alpha claim</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:t xml:space="preserve"> and more</w:t></w:r></w:p>
<w:p><w:r><w:t>Beta claim</w:t></w:r><w:r><w:endnoteReference w:id="3"/></w:r></w:p>
<w:p><w:r><w:t>Closing line</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/footnotes.xml',
		`<w:footnotes xmlns:w="${w}"><w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Footnote source</w:t></w:r></w:p></w:footnote></w:footnotes>`,
	);
	zip.file(
		'word/endnotes.xml',
		`<w:endnotes xmlns:w="${w}"><w:endnote w:id="3"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> Endnote source</w:t></w:r></w:p></w:endnote></w:endnotes>`,
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page): Promise<Locator> {
	page.on('dialog', (dialog) => void dialog.accept());
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'notes.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await fixture(),
	});
	const editor = page.locator('docx-editor');
	await expect(editor.locator('.ProseMirror').first()).toContainText('Closing line');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	await expect(editor.locator('.dve-print-footnote').first()).toBeVisible();
	return editor;
}

const noteEditor = (editor: Locator) => editor.locator('.dve-header-footer-editor .ProseMirror');

test('clicking a footnote on the page edits it in place, Escape returns to the body', async ({
	page,
}) => {
	const editor = await open(page);
	const footnote = editor.locator('.dve-print-footnote[data-note-id="2"]');
	await expect(footnote).toContainText('Footnote source');
	const text = footnote.locator('.dve-print-line span').last();
	const box = (await text.boundingBox())!;
	await page.mouse.click(box.x + box.width - 1, box.y + box.height / 2);
	await expect(noteEditor(editor)).toBeVisible();
	await page.keyboard.press('End');
	await page.keyboard.type(' edited');
	await expect(noteEditor(editor)).toContainText('Footnote source edited');
	await page.keyboard.press('Escape');
	await expect(noteEditor(editor)).toHaveCount(0);
	await expect(editor.locator('.dve-print-footnote[data-note-id="2"]')).toContainText(
		'Footnote source edited',
	);
	// The body takes typing again.
	const body = editor.locator('.dve-print-line', { hasText: 'Closing line' }).first();
	const closing = (await body.boundingBox())!;
	await page.mouse.click(closing.x + 2, closing.y + closing.height / 2);
	await page.keyboard.type('X');
	await expect(editor.locator('.ProseMirror').first()).toContainText('XClosing line');
	await expect(editor.locator('.dve-print-footnote[data-note-id="2"]')).toContainText(
		'Footnote source edited',
	);
});

test('clicking the body while a note is open closes it and places the caret there', async ({
	page,
}) => {
	const editor = await open(page);
	const footnote = editor.locator('.dve-print-footnote[data-note-id="2"]');
	const box = (await footnote.locator('.dve-print-line span').last().boundingBox())!;
	await page.mouse.click(box.x + box.width - 1, box.y + box.height / 2);
	await expect(noteEditor(editor)).toBeVisible();
	await page.keyboard.type('!');
	const closing = (await editor
		.locator('.dve-print-line', { hasText: 'Closing line' })
		.first()
		.boundingBox())!;
	await page.mouse.click(closing.x + 2, closing.y + closing.height / 2);
	await expect(noteEditor(editor)).toHaveCount(0);
	await page.keyboard.type('Y');
	await expect(editor.locator('.ProseMirror').first()).toContainText('YClosing line');
	await expect(editor.locator('.dve-print-footnote[data-note-id="2"]')).toContainText('!');
});

test('an endnote at the end of the document is editable on its page', async ({ page }) => {
	const editor = await open(page);
	const line = editor.locator('.dve-print-line', { hasText: 'Endnote source' }).first();
	const box = (await line.boundingBox())!;
	await page.mouse.click(box.x + box.width - 1, box.y + box.height / 2);
	await expect(noteEditor(editor)).toBeVisible();
	await page.keyboard.press('End');
	await page.keyboard.type('!');
	await expect(noteEditor(editor)).toContainText('Endnote source!');
	await page.keyboard.press('Escape');
	await expect(editor.locator('.dve-print-line', { hasText: 'Endnote source!' })).toHaveCount(1);
});

test('double-clicking a reference jumps to its note and the note number jumps back', async ({
	page,
}) => {
	const editor = await open(page);
	const reference = editor
		.locator('.dve-print-column .dve-print-line', { hasText: 'Alpha claim' })
		.first()
		.locator('span', { hasText: /^1$/ });
	await reference.dblclick();
	await expect(noteEditor(editor)).toBeVisible();
	await expect(noteEditor(editor)).toContainText('Footnote source');
	await page.keyboard.press('Escape');
	await expect(noteEditor(editor)).toHaveCount(0);
	const mark = editor.locator('.dve-print-footnote[data-note-id="2"] .dve-print-line span').first();
	await mark.dblclick();
	await expect(noteEditor(editor)).toHaveCount(0);
	// The caret sits right after the reference, so typing lands before " and more".
	await page.keyboard.type('Z');
	await expect(editor.locator('.ProseMirror').first()).toContainText('Z and more');
});

test('a note edit is saved into the note part and undone with the note editor', async ({
	page,
}) => {
	const editor = await open(page);
	const box = (await editor
		.locator('.dve-print-footnote[data-note-id="2"] .dve-print-line span')
		.last()
		.boundingBox())!;
	await page.mouse.click(box.x + box.width - 1, box.y + box.height / 2);
	await page.keyboard.press('End');
	await page.keyboard.type(' saved');
	await page.keyboard.press('Escape');
	const bytes = Buffer.from(
		await editor.evaluate(async (node) =>
			Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		),
	);
	const zip = await JSZip.loadAsync(bytes);
	expect(await zip.file('word/footnotes.xml')!.async('string')).toContain('Footnote source saved');
});
