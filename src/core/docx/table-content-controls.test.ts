import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index';
import { selectListItem } from './content-control-edit';
import type { Table } from './model';
import { WORD_NS } from './xml';

const hand = async (body: string) => {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${WORD_NS}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
};
const documentXml = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')) as string;
const tableOf = (blocks: unknown[]): Table => {
	const table = blocks.find((block) => (block as Table).type === 'table');
	if (!table) throw new Error('table expected');
	return table as Table;
};

const text = (value: string) => `<w:p><w:r><w:t xml:space="preserve">${value}</w:t></w:r></w:p>`;
const placeholderRun = (value: string) =>
	`<w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>${value}</w:t></w:r></w:p>`;
const choice = (id: string, extra = '') =>
	`<w:sdt><w:sdtPr><w:alias w:val="Choice ${id}"/><w:id w:val="${id}"/>${extra}<w:showingPlcHdr/>` +
	`<w:dropDownList><w:listItem w:displayText="Draft" w:value="draft"/>` +
	`<w:listItem w:displayText="Final" w:value="final"/></w:dropDownList></w:sdtPr>` +
	`<w:sdtContent><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${placeholderRun('Choose an item.')}</w:tc></w:sdtContent></w:sdt>`;
const grid = (rows: string) =>
	`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>${rows}</w:tbl>`;
const row = (cells: string) => `<w:tr>${cells}</w:tr>`;
const label = (value: string) => `<w:tc>${text(value)}</w:tc>`;

/** Picks entry `index` in the cell control of row `r`, column `c`, the way a user would. */
function pick(table: Table, r: number, c: number, index: number) {
	const cell = table.rows[r]![c]!;
	const change = selectListItem(cell.contentControl!, index)!;
	cell.contentControl = change.control;
	cell.paragraphs[0]!.runs = [{ text: change.text }];
}

describe('row and cell content controls on save', () => {
	it('clears the placeholder flag of a cell drop-down once a value is picked', async () => {
		const loaded = await hand(
			grid(row(`${label('Status')}${choice('11')}`) + row(`${label('Previous')}${choice('12')}`)),
		);
		pick(tableOf(loaded.model.blocks), 0, 1, 1);
		const xml = await documentXml(await loaded.save(loaded.model));
		const picked = xml.slice(xml.indexOf('Choice 11'), xml.indexOf('Choice 12'));
		expect(picked).not.toContain('showingPlcHdr');
		expect(picked).toContain('<w:t>Final</w:t>');
		const reopened = await loadDocx(await loaded.save(loaded.model));
		const control = tableOf(reopened.model.blocks).rows[0]![1]!.contentControl;
		expect(control).toMatchObject({ level: 'cell', kind: 'dropDownList', alias: 'Choice 11' });
		expect(control?.showingPlaceholder).toBeUndefined();
	});

	it('leaves the markup of the cell controls it did not change as it was', async () => {
		const source = choice('12', '<w:tag w:val="previous"/>');
		const loaded = await hand(
			grid(row(`${label('Status')}${choice('11')}`) + row(`${label('Previous')}${source}`)),
		);
		pick(tableOf(loaded.model.blocks), 0, 1, 0);
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toContain(source);
	});

	it('saves an edited row control without adding row properties', async () => {
		const loaded = await hand(
			grid(
				`<w:sdt><w:sdtPr><w:alias w:val="Entry"/><w:id w:val="21"/></w:sdtPr><w:sdtContent>` +
					`${row(`${label('Name')}${label('Value')}`)}</w:sdtContent></w:sdt>`,
			),
		);
		const table = tableOf(loaded.model.blocks);
		table.rowProperties![0]!.contentControl = {
			...table.rowProperties![0]!.contentControl!,
			alias: 'Entry 1',
			tag: 'entry',
		};
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toContain(
			'<w:sdtPr><w:alias w:val="Entry 1"/><w:tag w:val="entry"/><w:id w:val="21"/></w:sdtPr>',
		);
		expect(xml).not.toContain('<w:trPr');
	});

	it('takes away the control of a cell the model no longer gives one, and keeps the cell', async () => {
		const loaded = await hand(grid(row(`${label('Status')}${choice('11')}`)));
		const table = tableOf(loaded.model.blocks);
		delete table.rows[0]![1]!.contentControl;
		table.rows[0]![1]!.paragraphs[0]!.runs = [{ text: 'Final' }];
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).not.toContain('<w:sdt>');
		expect(xml).toContain(
			'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Final</w:t></w:r></w:p></w:tc></w:tr>',
		);
	});

	it('takes away a control around several cells or rows once none of them has it', async () => {
		const wrap = (content: string) =>
			`<w:sdt><w:sdtPr><w:alias w:val="Group"/><w:id w:val="21"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
		const cells = await hand(grid(row(wrap(`${label('Left')}${label('Right')}`))));
		const pair = tableOf(cells.model.blocks).rows[0]!;
		expect(pair.map((cell) => cell.contentControl?.id)).toEqual(['21', '21']);
		for (const cell of pair) delete cell.contentControl;
		const cellXml = await documentXml(await cells.save(cells.model));
		expect(cellXml).not.toContain('<w:sdt>');
		expect(cellXml).toContain(`<w:tr>${label('Left')}${label('Right')}</w:tr>`);
		const rows = await hand(
			grid(wrap(row(`${label('A')}${label('B')}`) + row(`${label('C')}${label('D')}`))),
		);
		const table = tableOf(rows.model.blocks);
		expect(table.rowProperties?.map((props) => props.contentControl?.id)).toEqual(['21', '21']);
		for (const props of table.rowProperties ?? []) delete props.contentControl;
		const rowXml = await documentXml(await rows.save(rows.model));
		expect(rowXml).not.toContain('<w:sdt>');
		expect(rowXml).toContain(`</w:tblGrid>${row(`${label('A')}${label('B')}`)}<w:tr>`);
	});

	it('keeps the range markers Word writes beside a cell or row control when unwrapping it', async () => {
		// Word allows bookmarks, comment ranges and move ranges as children of a row or cell w:sdt
		// (MS-OI29500 2.1.202 and 2.1.204), next to its w:sdtContent.
		const marked = (content: string) =>
			`<w:sdt><w:sdtPr><w:id w:val="22"/><w:temporary/></w:sdtPr><w:bookmarkStart w:id="7" w:name="Kept"/>` +
			`<w:sdtContent>${content}</w:sdtContent><w:bookmarkEnd w:id="7"/></w:sdt>`;
		const loaded = await hand(grid(row(`${label('Status')}${marked(label('Open'))}`)));
		delete tableOf(loaded.model.blocks).rows[0]![1]!.contentControl;
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).not.toContain('<w:sdt>');
		expect(xml).toContain(
			`${label('Status')}<w:bookmarkStart w:id="7" w:name="Kept"/>${label('Open')}<w:bookmarkEnd w:id="7"/></w:tr>`,
		);
		const rows = await hand(grid(marked(row(`${label('A')}${label('B')}`))));
		delete tableOf(rows.model.blocks).rowProperties![0]!.contentControl;
		const rowXml = await documentXml(await rows.save(rows.model));
		expect(rowXml).toContain(
			`</w:tblGrid><w:bookmarkStart w:id="7" w:name="Kept"/>${row(`${label('A')}${label('B')}`)}<w:bookmarkEnd w:id="7"/>`,
		);
	});

	it('writes the row and cell controls of a table added to the document', async () => {
		const loaded = await hand(text('Before'));
		loaded.model.blocks.push({
			type: 'table',
			id: 'added',
			rows: [
				[
					{ paragraphs: [{ type: 'paragraph', id: 'added-a', runs: [{ text: 'Status' }] }] },
					{
						paragraphs: [{ type: 'paragraph', id: 'added-b', runs: [{ text: 'Draft' }] }],
						contentControl: { id: '41', level: 'cell', kind: 'richText', alias: 'Status' },
					},
				],
			],
			rowProperties: [{ contentControl: { id: '42', level: 'row', kind: 'richText' } }],
		});
		const reopened = await loadDocx(await loaded.save(loaded.model));
		const table = tableOf(reopened.model.blocks);
		expect(table.rowProperties?.[0]?.contentControl).toMatchObject({ id: '42', level: 'row' });
		expect(table.rows[0]![1]!.contentControl).toMatchObject({ id: '41', alias: 'Status' });
		expect(table.rows[0]![1]!.paragraphs[0]!.runs[0]!.text).toBe('Draft');
	});

	it('saves a copied cell control with the new id it was given', async () => {
		const loaded = await hand(
			grid(row(`${label('Status')}${choice('11')}`) + row(`${label('Again')}${label('Draft')}`)),
		);
		const table = tableOf(loaded.model.blocks);
		// A pasted copy keeps the properties it was copied with and takes a new id.
		table.rows[1]![1]!.contentControl = { ...table.rows[0]![1]!.contentControl!, id: '9001' };
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml.match(/<w:id w:val="11"\/>/g)).toHaveLength(1);
		expect(xml).toContain('<w:alias w:val="Choice 11"/><w:id w:val="9001"/>');
	});

	it('wraps a cell in the control the model gives it', async () => {
		const loaded = await hand(grid(row(`${label('Remarks')}${label('None')}`)));
		tableOf(loaded.model.blocks).rows[0]![1]!.contentControl = {
			id: '31',
			level: 'cell',
			kind: 'richText',
			alias: 'Remarks',
		};
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toContain(
			'<w:sdt><w:sdtPr><w:alias w:val="Remarks"/><w:id w:val="31"/></w:sdtPr><w:sdtContent><w:tc>',
		);
		const reopened = await loadDocx(await loaded.save(loaded.model));
		expect(tableOf(reopened.model.blocks).rows[0]![1]!.contentControl).toMatchObject({
			id: '31',
			level: 'cell',
			alias: 'Remarks',
		});
	});
});
