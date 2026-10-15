import { expect, test, type Locator, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';

test.use({ viewport: { width: 1366, height: 1500 } });

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const borders =
	'<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders>';
const para = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (content: string, width = 3000) =>
	`<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
const nested = `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/>${borders}</w:tblPr><w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>
<w:tr>${cell(para('N11'), 2500)}${cell(para('N12'), 2500)}</w:tr>
<w:tr>${cell(para('N21'), 2500)}${cell(para('N22'), 2500)}</w:tr></w:tbl>`;
const outer = `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>${borders}</w:tblPr><w:tblGrid><w:gridCol w:w="6000"/><w:gridCol w:w="3000"/></w:tblGrid>
<w:tr>${cell(para('Outer') + nested + para('Tail'), 6000)}${cell(para('Next'))}</w:tr></w:tbl>`;

async function docx(body: string): Promise<Buffer> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body>${body}<w:p><w:r><w:t>After</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page): Promise<Locator> {
	page.on('dialog', (dialog) => void dialog.accept());
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'nested.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await docx(outer),
	});
	const editor = page.locator('docx-editor');
	await expect(editor.locator('.ProseMirror').first()).toContainText('After');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	await expect(editor.locator('.dve-print-line').first()).toBeVisible();
	return editor;
}

const lineOf = (editor: Locator, text: string) =>
	editor.locator('.dve-print-line', { hasText: text }).first();

async function clickAtStart(page: Page, editor: Locator, text: string) {
	const box = (await lineOf(editor, text).boundingBox())!;
	await page.mouse.click(box.x + 2, box.y + box.height / 2);
}

const input = (editor: Locator) => editor.locator('.ProseMirror').first();

test('nested table paragraphs are editable on the page', async ({ page }) => {
	const editor = await open(page);
	await expect(editor.locator('.dve-print-block[data-nested="true"]').first()).toBeVisible();
	await expect(editor.locator('[data-read-only="true"]')).toHaveCount(0);
	await clickAtStart(page, editor, 'N12');
	await page.keyboard.press('End');
	await page.keyboard.type('!');
	await expect(input(editor).locator('table table td', { hasText: 'N12!' })).toHaveCount(1);
	await expect(lineOf(editor, 'N12!')).toBeVisible();
	await page.keyboard.press('Control+z');
	await expect(lineOf(editor, 'N12')).toHaveText('N12');
});

test('Up and Down step between nested rows under the same column', async ({ page }) => {
	const editor = await open(page);
	await clickAtStart(page, editor, 'N12');
	await page.keyboard.press('ArrowDown');
	await page.keyboard.type('D');
	await expect(input(editor).locator('table table td', { hasText: 'DN22' })).toHaveCount(1);
	await page.keyboard.press('ArrowUp');
	await page.keyboard.type('U');
	await expect(input(editor).locator('table table td').nth(1)).toContainText('U');
	await page.keyboard.press('ArrowUp');
	await page.keyboard.type('O');
	await expect(input(editor).locator('td').first().locator('> p').first()).toContainText('O');
});

test('Tab walks nested cells then continues in the enclosing table', async ({ page }) => {
	const editor = await open(page);
	await clickAtStart(page, editor, 'N11');
	await page.keyboard.press('Tab');
	await page.keyboard.type('a');
	await page.keyboard.press('Tab');
	await page.keyboard.type('b');
	await page.keyboard.press('Tab');
	await page.keyboard.type('c');
	await page.keyboard.press('Tab');
	await page.keyboard.type('d');
	const cells = await input(editor).locator('table table td').allInnerTexts();
	expect(cells.map((t) => t.trim())).toEqual(['N11', 'a', 'b', 'c']);
	await expect(input(editor).locator('td', { hasText: 'd' }).last()).toContainText('d');
	await page.keyboard.press('Shift+Tab');
	await page.keyboard.type('e');
	await expect(input(editor).locator('table table td').nth(3)).toContainText('e');
});

test('an edit in a nested cell is saved into the nested table', async ({ page }) => {
	const editor = await open(page);
	await clickAtStart(page, editor, 'N21');
	await page.keyboard.press('End');
	await page.keyboard.type('X');
	const bytes = Buffer.from(
		await editor.evaluate(async (node) =>
			Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		),
	);
	const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
	expect(xml).toContain('N21X');
	expect(xml.match(/<w:tbl>/g)).toHaveLength(2);
	expect(xml).toContain('Outer');
	expect(xml).toContain('Tail');
});
