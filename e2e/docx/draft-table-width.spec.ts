import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** A 12-column auto-fit table and a centred one with thick borders and large cell margins (#322). */
async function fixture() {
	const zip = new JSZip();
	const cell = (text: string, width: number) =>
		`<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
	const columns = Array.from({ length: 12 }, (_, i) => `Colonne${i + 1}`);
	const wide = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${columns.map(() => '<w:gridCol w:w="780"/>').join('')}</w:tblGrid><w:tr>${columns.map((text) => cell(text, 780)).join('')}</w:tr></w:tbl>`;
	const line = 'w:val="single" w:sz="24" w:space="0" w:color="000000"';
	const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
		.map((side) => `<w:${side} ${line}/>`)
		.join('');
	const centred = `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders>${borders}</w:tblBorders><w:tblCellMar><w:left w:w="360" w:type="dxa"/><w:right w:w="360" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid><w:tr>${cell('https://example.com/a/very/long/unbroken/address/that/does/not/fit', 4680)}${cell('B', 4680)}</w:tr></w:tbl>`;
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${wide}<w:p/>${centred}<w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}

test('Draft view keeps tables inside the text column, as Print Layout does (#322)', async ({
	page,
}) => {
	page.on('dialog', (dialog) => void dialog.accept());
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'draft-table-width.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await fixture(),
	});
	const editor = page.locator('docx-editor');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('combobox', { name: 'Layout view', exact: true }).selectOption('draft');
	const tables = editor.locator('.ProseMirror table');
	await expect(tables).toHaveCount(2);
	const paper = await editor.locator('.dve-paper').evaluate((element) => {
		const box = element.getBoundingClientRect();
		const style = getComputedStyle(element);
		return {
			left: box.left + parseFloat(style.paddingLeft),
			right: box.right - parseFloat(style.paddingRight),
		};
	});
	for (const table of await tables.all()) {
		const box = (await table.boundingBox())!;
		// Half of a collapsed outer border may sit outside the text column, as in Word.
		expect(box.x).toBeGreaterThanOrEqual(paper.left - 4);
		expect(box.x + box.width).toBeLessThanOrEqual(paper.right + 4);
	}
});
