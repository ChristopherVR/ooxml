import { expect, test, type Locator, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';

test.use({ viewport: { width: 1366, height: 900 } });

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

async function longDocument(): Promise<Buffer> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	);
	zip.file(
		'word/_rels/document.xml.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>',
	);
	zip.file(
		'word/header1.xml',
		`<w:hdr xmlns:w="${w}"><w:p><w:r><w:t>Running head</w:t></w:r></w:p></w:hdr>`,
	);
	zip.file(
		'word/footnotes.xml',
		`<w:footnotes xmlns:w="${w}"><w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> First note</w:t></w:r></w:p></w:footnote></w:footnotes>`,
	);
	const pages = Array.from(
		{ length: 14 },
		(_, index) =>
			`<w:p><w:r><w:t>Page ${index + 1}</w:t></w:r>${index === 0 ? '<w:r><w:footnoteReference w:id="2"/></w:r>' : ''}</w:p>${index < 13 ? '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' : ''}`,
	).join('');
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body>${pages}<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}

/** Puts the body caret on a late page, so the caret's own page is not what keeps page 1 mounted. */
async function caretOnLastPages(page: Page, editor: Locator): Promise<void> {
	const last = editor.locator('.dve-print-page').nth(12);
	await last.scrollIntoViewIfNeeded();
	const line = editor.locator('.dve-print-line', { hasText: 'Page 13' }).first();
	await line.scrollIntoViewIfNeeded();
	const box = (await line.boundingBox())!;
	await page.mouse.click(box.x + 2, box.y + box.height / 2);
	await page.keyboard.type('Q');
	await expect(editor.locator('.ProseMirror').first()).toContainText('QPage 13');
	// Let the debounced relayout redraw the pages before scrolling around.
	await expect(editor.locator('.dve-print-line', { hasText: 'QPage 13' })).toHaveCount(1);
	await editor.locator('.dve-print-page').first().scrollIntoViewIfNeeded();
	await expect(editor.locator('.dve-print-page').first()).not.toBeEmpty();
}

async function open(page: Page): Promise<Locator> {
	page.on('dialog', (dialog) => void dialog.accept());
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'long.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await longDocument(),
	});
	const editor = page.locator('docx-editor');
	await expect(editor.locator('.ProseMirror').first()).toContainText('Page 14');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	await expect(editor.locator('.dve-print-page').first()).toBeVisible();
	return editor;
}

test('an open header editor survives scrolling away from its page and back', async ({ page }) => {
	const editor = await open(page);
	await expect(editor.locator('.dve-print-page')).toHaveCount(14);
	await caretOnLastPages(page, editor);
	const header = editor.locator('.dve-print-page').first().locator('.dve-print-header');
	await expect(header).toContainText('Running head');
	await header.scrollIntoViewIfNeeded();
	const headerBox = (await header.boundingBox())!;
	await page.mouse.dblclick(headerBox.x + 20, headerBox.y + headerBox.height / 2);
	await expect(editor.locator('.dve-header-footer-editor .ProseMirror')).toBeVisible();
	await page.keyboard.type('One ');
	// Page 10 is far outside the render window around page 1.
	await editor.locator('.dve-print-page').nth(9).scrollIntoViewIfNeeded();
	await expect(editor.locator('.dve-print-page').nth(9)).not.toBeEmpty();
	await expect(editor.locator('.dve-header-footer-editor .ProseMirror')).toBeVisible();
	await expect(editor.locator('.dve-print-page').nth(1)).toBeEmpty();
	await editor.locator('.dve-print-page').first().scrollIntoViewIfNeeded();
	await page.keyboard.type('Two');
	await expect(editor.locator('.dve-header-footer-editor .ProseMirror')).toContainText('One Two');
});

test('an open note editor survives scrolling away from its page', async ({ page }) => {
	const editor = await open(page);
	await caretOnLastPages(page, editor);
	const note = editor.locator('.dve-print-footnote[data-note-id="2"]');
	await note.scrollIntoViewIfNeeded();
	const box = (await note.locator('.dve-print-line span').last().boundingBox())!;
	await page.mouse.click(box.x + box.width - 1, box.y + box.height / 2);
	await expect(editor.locator('.dve-header-footer-editor .ProseMirror')).toBeVisible();
	await editor.locator('.dve-print-page').nth(10).scrollIntoViewIfNeeded();
	await expect(editor.locator('.dve-print-page').nth(10)).not.toBeEmpty();
	await expect(editor.locator('.dve-header-footer-editor .ProseMirror')).toBeVisible();
	await expect(editor.locator('.dve-print-page').first()).not.toBeEmpty();
});
