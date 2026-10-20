import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index';
import { checkboxState } from './content-control-edit';
import { at, expectTable, must } from './test-support/access';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const w14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const text = (value: string) => `<w:p><w:r><w:t xml:space="preserve">${value}</w:t></w:r></w:p>`;
const cell = (content: string, props = '') =>
	`<w:tc>${props ? `<w:tcPr>${props}</w:tcPr>` : ''}${content}</w:tc>`;
const label = (value: string, props = '') =>
	cell(text(value), `<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>${props}`);
const span = (columns: number) => `<w:gridSpan w:val="${columns}"/>`;
const checkbox =
	`<w:sdt><w:sdtPr><w:id w:val="7"/><w14:checkbox><w14:checked w14:val="0"/>` +
	`<w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>` +
	`</w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic"/></w:rPr>` +
	`<w:t>☐</w:t></w:r></w:sdtContent></w:sdt>`;

/** A label/value form grid: four grid columns, rows of 2, 4, 1, 2 and 2 physical cells. */
const formGridXml = `<w:document xmlns:w="${ns}" xmlns:w14="${w14}"><w:body><w:tbl>
<w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2500"/><w:gridCol w:w="1500"/><w:gridCol w:w="3000"/></w:tblGrid>
<w:tr>${label('Document number')}${cell(text('Number'), span(3))}</w:tr>
<w:tr>${label('Revision')}${cell(text('1'))}${label('Date')}${cell(text('Date'))}</w:tr>
<w:tr>${cell(text('Change history'), span(4))}</w:tr>
<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>${cell(text('Indented'))}${cell(text('Row'), span(2))}</w:tr>
<w:tr>${label('Approved')}${cell(`<w:p>${checkbox}<w:r><w:t xml:space="preserve"> Yes</w:t></w:r></w:p>`, span(3))}</w:tr>
</w:tbl><w:sectPr/></w:body></w:document>`;

async function loadGrid() {
	const zip = new JSZip();
	zip.file('word/document.xml', formGridXml);
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	return { loaded, table: expectTable(loaded.model.blocks[0]) };
}

async function documentXml(bytes: Uint8Array): Promise<string> {
	return (await (await JSZip.loadAsync(bytes)).file('word/document.xml')?.async('string')) ?? '';
}

const rowsOf = (xml: string) => [...xml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map((m) => m[0]);
const cellCounts = (xml: string) => rowsOf(xml).map((row) => (row.match(/<w:tc\b/g) ?? []).length);

describe('DOCX text edits in tables whose rows have different cell counts', () => {
	it('reads the grid as it is: one model cell per physical cell, with its span', async () => {
		const { table } = await loadGrid();
		expect(table.rows.map((row) => row.length)).toEqual([2, 4, 1, 2, 2]);
		expect(at(at(table.rows, 0), 1).gridSpan).toBe(3);
		expect(table.structureEditable).toBe(false);
	});

	it('saves a value typed into a merged cell and keeps every other cell untouched', async () => {
		const { loaded, table } = await loadGrid();
		const value = at(at(at(table.rows, 0), 1).paragraphs, 0);
		at(value.runs, 0).text = 'DOC-0042';
		const xml = await documentXml(await loaded.save());
		expect(cellCounts(xml)).toEqual([2, 4, 1, 2, 2]);
		const [first, second, third, fourth] = rowsOf(xml);
		expect(first).toContain('<w:gridSpan w:val="3"/>');
		expect(first).toContain('DOC-0042');
		expect(first).not.toContain('>Number<');
		expect(first).toContain('w:fill="D9D9D9"');
		// Rows without edits keep their source XML exactly.
		expect(second).toBe(rowsOf(formGridXml)[1]);
		expect(third).toBe(rowsOf(formGridXml)[2]);
		expect(fourth).toBe(rowsOf(formGridXml)[3]);
	});

	it('saves edits in rows that follow narrower and wider rows, and after gridBefore', async () => {
		const { loaded, table } = await loadGrid();
		at(at(at(at(table.rows, 1), 3).paragraphs, 0).runs, 0).text = '20-Oct-2026';
		at(at(at(at(table.rows, 2), 0).paragraphs, 0).runs, 0).text = 'History';
		at(at(at(at(table.rows, 3), 1).paragraphs, 0).runs, 0).text = 'Second cell';
		const saved = await loaded.save();
		const xml = await documentXml(saved);
		expect(cellCounts(xml)).toEqual([2, 4, 1, 2, 2]);
		expect(at(rowsOf(xml), 3)).toContain('<w:gridBefore w:val="1"/>');
		const reread = expectTable((await loadDocx(saved)).model.blocks[0]);
		const texts = reread.rows.map((row) =>
			row.map((item) => item.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('')),
		);
		expect(texts).toEqual([
			['Document number', 'Number'],
			['Revision', '1', 'Date', '20-Oct-2026'],
			['History'],
			['Indented', 'Second cell'],
			['Approved', '☐ Yes'],
		]);
	});

	it('saves a check box toggled inside the grid', async () => {
		const { loaded, table } = await loadGrid();
		const paragraph = at(at(at(table.rows, 4), 1).paragraphs, 0);
		const run = must(paragraph.runs.find((item) => item.contentControl?.kind === 'checkbox'));
		const change = checkboxState(must(run.contentControl), true);
		run.contentControl = change.control;
		run.text = change.text;
		const xml = await documentXml(await loaded.save());
		expect(xml).toContain('<w14:checked w14:val="1"/>');
		expect(xml).toContain('<w:t>☒</w:t>');
		expect(cellCounts(xml)).toEqual([2, 4, 1, 2, 2]);
	});

	it('still refuses structural edits, including emptying a cell', async () => {
		const added = await loadGrid();
		added.table.rows.push([
			{ paragraphs: [{ type: 'paragraph', id: 'added-left', runs: [{ text: '' }] }] },
			{ paragraphs: [{ type: 'paragraph', id: 'added-right', runs: [{ text: '' }] }] },
		]);
		await expect(added.loaded.save()).rejects.toThrow('Cannot change the structure');

		const emptied = await loadGrid();
		at(at(emptied.table.rows, 1), 1).paragraphs = [];
		await expect(emptied.loaded.save()).rejects.toThrow('Cannot change the structure');
	});
});
