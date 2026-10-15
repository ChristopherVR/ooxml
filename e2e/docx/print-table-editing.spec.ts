import { expect, test, type Locator, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';

test.use({ viewport: { width: 1366, height: 1500 } });

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const cell = (text: string) =>
	`<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const borders =
	'<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders>';
const table = (rows: string[][], rowHeight = 0) =>
	`<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>${borders}</w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>${rows
		.map(
			(row) =>
				`<w:tr>${rowHeight ? `<w:trPr><w:trHeight w:val="${rowHeight}" w:hRule="atLeast"/></w:trPr>` : ''}${row.map(cell).join('')}</w:tr>`,
		)
		.join('')}</w:tbl>`;

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

async function open(page: Page, body: string): Promise<Locator> {
	page.on('dialog', (dialog) => void dialog.accept());
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'cells.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await docx(body),
	});
	const editor = page.locator('docx-editor');
	await expect(editor.locator('.ProseMirror').first()).toContainText('After');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	await expect(editor.locator('.dve-print-line').first()).toBeVisible();
	return editor;
}

const grid = [
	['A1', 'B1', 'C1'],
	['A2', 'B2', 'C2'],
	['A3', 'B3', 'C3'],
];
const cellsText = (editor: Locator) =>
	editor.locator('.ProseMirror').first().locator('td, th').allInnerTexts();

test('click in a cell and type; the cell is not read-only on the page', async ({ page }) => {
	const editor = await open(page, table(grid));
	const line = editor.locator('.dve-print-line', { hasText: 'B2' }).first();
	await expect(line.locator('xpath=ancestor::*[@data-read-only="true"]')).toHaveCount(0);
	const box = (await line.boundingBox())!;
	await page.mouse.click(box.x + 2, box.y + box.height / 2);
	await page.keyboard.press('End');
	await page.keyboard.type('!');
	expect((await cellsText(editor)).map((t) => t.trim())).toContain('B2!');
});

test('Tab and Shift+Tab move between cells, Up and Down between rows', async ({ page }) => {
	const editor = await open(page, table(grid));
	const a1 = (await editor.locator('.dve-print-line', { hasText: 'A1' }).first().boundingBox())!;
	await page.mouse.click(a1.x + 2, a1.y + a1.height / 2);
	// Tab selects the next cell's text, as Word does, so typing replaces it.
	await page.keyboard.press('Tab');
	await page.keyboard.type('1');
	await page.keyboard.press('Tab');
	await page.keyboard.type('2');
	await page.keyboard.press('Shift+Tab');
	await page.keyboard.type('3');
	let texts = (await cellsText(editor)).map((t) => t.trim());
	expect(texts.slice(0, 3)).toEqual(['A1', '3', '2']);
	await page.keyboard.press('ArrowDown');
	await page.keyboard.type('D');
	await page.keyboard.press('ArrowDown');
	await page.keyboard.type('E');
	await page.keyboard.press('ArrowUp');
	await page.keyboard.press('ArrowUp');
	await page.keyboard.type('U');
	texts = (await cellsText(editor)).map((t) => t.trim());
	expect(texts[1]).toContain('U');
	expect(texts[4]).toContain('D');
	expect(texts[7]).toContain('E');
});

test('Tab in the last cell adds a row', async ({ page }) => {
	const editor = await open(page, table(grid));
	const c3 = (await editor.locator('.dve-print-line', { hasText: 'C3' }).first().boundingBox())!;
	await page.mouse.click(c3.x + 2, c3.y + c3.height / 2);
	await page.keyboard.press('End');
	await page.keyboard.press('Tab');
	await page.keyboard.type('N');
	await expect(editor.locator('.ProseMirror').first().locator('tr')).toHaveCount(4);
	await expect(editor.locator('.dve-print-line', { hasText: 'N' }).first()).toBeVisible();
});

test('a table split across two pages keeps editing and arrowing', async ({ page }) => {
	const rows = Array.from({ length: 14 }, (_, i) => [`A${i}`, `B${i}`, `C${i}`]);
	const editor = await open(page, table(rows, 1200));
	await expect(editor.locator('.dve-print-page')).not.toHaveCount(1);
	const last = editor.locator('.dve-print-line', { hasText: 'B13' }).first();
	await last.scrollIntoViewIfNeeded();
	const box = (await last.boundingBox())!;
	await page.mouse.click(box.x + 2, box.y + box.height / 2);
	await page.keyboard.type('Z');
	await page.keyboard.press('ArrowUp');
	await page.keyboard.type('Y');
	const texts = (await cellsText(editor)).map((t) => t.trim());
	expect(texts).toContain('ZB13');
	expect(texts.some((t) => t.includes('Y') && t.includes('12'))).toBe(true);
	// Up across the page break reaches earlier rows.
	for (let i = 0; i < 8; i++) await page.keyboard.press('ArrowUp');
	await page.keyboard.type('X');
	expect((await cellsText(editor)).some((t) => t.includes('X'))).toBe(true);
});
