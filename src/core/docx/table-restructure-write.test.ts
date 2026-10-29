// Saving restructured merged and nested tables: the edits come from the core structure engine,
// the XML from the merge-aware writer. Unknown tblPr, trPr and tcPr children must survive, and the
// saved package must read back to the grid the engine produced.
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index';
import type { Table, TableCell } from './model';
import { tableGrid } from './table-grid';
import { deleteColumns, deleteRow, insertColumn, insertRow } from './table-structure';
import { mergeCells, splitCell } from './table-structure-merge';
import { expectTable, must } from './test-support/access';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const w14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
let serial = 0;
const newId = () => `fresh-${++serial}`;
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const tc = (text: string, props = '') =>
	`<w:tc><w:tcPr>${props}</w:tcPr>${text ? p(text) : '<w:p/>'}</w:tc>`;
const tcW = (w: number) => `<w:tcW w:w="${w}" w:type="dxa"/>`;
const span = (n: number) => `<w:gridSpan w:val="${n}"/>`;
const NOWRAP = '<w:noWrap/>';

/** Three grid columns: a horizontal merge in row 1, a vertical merge in column 3, unknown properties throughout. */
const MERGED =
	`<w:document xmlns:w="${ns}" xmlns:w14="${w14}"><w:body><w:tbl>` +
	`<w:tblPr><w:tblStyle w:val="Grid"/><w:tblW w:w="9000" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCaption w:val="Kept caption"/></w:tblPr>` +
	`<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
	`<w:tr w14:paraId="0000000A"><w:trPr><w:cantSplit/><w:hidden/></w:trPr>${tc('A', tcW(6000) + span(2) + NOWRAP)}${tc('V', tcW(3000) + '<w:vMerge w:val="restart"/>')}</w:tr>` +
	`<w:tr><w:trPr><w:trHeight w:val="500" w:hRule="atLeast"/></w:trPr>${tc('D', tcW(3000) + NOWRAP)}${tc('E', tcW(3000))}${tc('', tcW(3000) + '<w:vMerge/>')}</w:tr>` +
	`<w:tr>${tc('G', tcW(3000))}${tc('H', tcW(3000) + NOWRAP)}${tc('', tcW(3000) + '<w:vMerge/>')}</w:tr>` +
	`</w:tbl><w:sectPr/></w:body></w:document>`;

const SIMPLE =
	`<w:document xmlns:w="${ns}"><w:body><w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr>` +
	`<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
	`<w:tr>${tc('a', tcW(3000) + NOWRAP)}${tc('b', tcW(3000))}</w:tr><w:tr>${tc('c', tcW(3000))}${tc('d', tcW(3000))}</w:tr>` +
	`</w:tbl><w:sectPr/></w:body></w:document>`;

async function open(xml: string) {
	const zip = new JSZip();
	zip.file('word/document.xml', xml);
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	return { loaded, table: expectTable(loaded.model.blocks[0]) };
}

/** Puts an edit's result into the model the way the editor does, then saves. */
async function apply(
	loaded: Awaited<ReturnType<typeof open>>['loaded'],
	table: Table,
	result: Table | undefined,
) {
	const edited = must(result, 'the edit result');
	table.rows = edited.rows;
	if (edited.grid) table.grid = edited.grid;
	if (edited.rowProperties) table.rowProperties = edited.rowProperties;
	else delete table.rowProperties;
	table.structureEditable = false;
	return loaded.save();
}

const xmlOf = async (bytes: Uint8Array) =>
	must(await (await JSZip.loadAsync(bytes)).file('word/document.xml')?.async('string'));
const rows = (xml: string) => [...xml.matchAll(/<w:tr[ >][\s\S]*?<\/w:tr>/g)].map((m) => m[0]);
const describe1 = (cell: TableCell) =>
	`${cell.paragraphs
		.map((q) => q.runs.map((r) => r.text).join(''))
		.filter(Boolean)
		.join('/')}${cell.gridSpan && cell.gridSpan > 1 ? `:${cell.gridSpan}` : ''}${
		cell.verticalMerge === 'restart' ? '^' : cell.verticalMerge === 'continue' ? '|' : ''
	}`;
const layoutOf = (table: Table) => table.rows.map((row) => row.map(describe1));
async function reread(bytes: Uint8Array) {
	const table = expectTable((await loadDocx(bytes)).model.blocks[0]);
	expect(tableGrid(table), 'saved table tiles its grid').toBeDefined();
	return table;
}

describe('saving restructured merged tables', () => {
	it('is restructurable: merged, with a consistent grid and nothing the engine cannot keep', async () => {
		const { table } = await open(MERGED);
		expect(table.structureEditable).toBe(false);
		expect(table.restructurable).toBe(true);
		expect(layoutOf(table)).toEqual([
			['A:2', 'V^'],
			['D', 'E', '|'],
			['G', 'H', '|'],
		]);
	});

	it('inserts a row through the vertical merge, keeping unknown properties of every surviving row and cell', async () => {
		const { loaded, table } = await open(MERGED);
		const saved = await apply(loaded, table, insertRow(table, 1, 'below', newId));
		const xml = await xmlOf(saved);
		const [first, second, added, last] = rows(xml);
		expect(rows(xml)).toHaveLength(4);
		expect(first).toContain('w14:paraId="0000000A"');
		expect(first).toContain('<w:hidden/>');
		expect(first).toContain(NOWRAP);
		expect(second).toContain('w:val="500"');
		// The new row copies the height, and its third cell continues the merge.
		expect(added).toContain('w:val="500"');
		expect(added).toMatch(/<w:vMerge( w:val="continue")?\/>/);
		expect(last).toContain('>G<');
		expect(xml).toContain('<w:tblCaption w:val="Kept caption"/>');
		expect(layoutOf(await reread(saved))).toEqual([
			['A:2', 'V^'],
			['D', 'E', '|'],
			['', '', '|'],
			['G', 'H', '|'],
		]);
	});

	it('deletes the row that starts a vertical merge and promotes the next row', async () => {
		const { loaded, table } = await open(MERGED);
		const saved = await apply(loaded, table, deleteRow(table, 0));
		const xml = await xmlOf(saved);
		expect(rows(xml)).toHaveLength(2);
		expect(xml).not.toContain('<w:hidden/>');
		expect(rows(xml)[0]).toContain('w:val="500"');
		expect(layoutOf(await reread(saved))).toEqual([
			['D', 'E', '^'],
			['G', 'H', '|'],
		]);
	});

	it('inserts a column: new grid column and cells, no change to the other widths', async () => {
		const { loaded, table } = await open(MERGED);
		const saved = await apply(loaded, table, insertColumn(table, 2, 'right', newId));
		const xml = await xmlOf(saved);
		expect(xml.match(/<w:gridCol /g)).toHaveLength(4);
		expect(xml).toContain(NOWRAP);
		const reopened = await reread(saved);
		expect(reopened.grid).toEqual([3000, 3000, 3000, 3000]);
		expect(layoutOf(reopened)).toEqual([
			['A:2', 'V^', ''],
			['D', 'E', '|', ''],
			['G', 'H', '|', ''],
		]);
	});

	it('deletes a column whose cells include a vertical merge', async () => {
		const { loaded, table } = await open(MERGED);
		const saved = await apply(loaded, table, deleteColumns(table, 2, 2));
		const reopened = await reread(saved);
		expect(reopened.grid).toEqual([3000, 3000]);
		expect(layoutOf(reopened)).toEqual([['A:2'], ['D', 'E'], ['G', 'H']]);
		expect(await xmlOf(saved)).not.toContain('w:vMerge');
	});

	it('merges cells across rows and splits them again', async () => {
		const { loaded, table } = await open(MERGED);
		const merged = must(
			mergeCells(table, { firstRow: 1, lastRow: 2, firstCol: 0, lastCol: 1 }, newId),
		);
		const saved = await apply(loaded, table, merged);
		const reopened = await reread(saved);
		expect(layoutOf(reopened)).toEqual([
			['A:2', 'V^'],
			['D/E/G/H:2^', '|'],
			[':2|', '|'],
		]);
		const xml = await xmlOf(saved);
		expect(rows(xml)[1]).toContain(NOWRAP);
		const again = await open(xml);
		const split = splitCell(again.table, 1, 0, 2, 2, newId);
		const savedAgain = await apply(again.loaded, again.table, split);
		expect((await reread(savedAgain)).rows).toHaveLength(3);
	});

	it('writes a split of a cell with the grid and spans the engine computed', async () => {
		const { loaded, table } = await open(MERGED);
		const saved = await apply(loaded, table, splitCell(table, 1, 1, 1, 2, newId));
		const reopened = await reread(saved);
		expect(reopened.grid?.length).toBe(4);
		expect(layoutOf(reopened)[1]).toEqual(['D', 'E', '', '|']);
	});

	it('turns a plain table into a merged one when two cells are merged', async () => {
		const { loaded, table } = await open(SIMPLE);
		expect(table.structureEditable).toBe(true);
		const merged = must(
			mergeCells(table, { firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 1 }, newId),
		);
		const saved = await apply(loaded, table, merged);
		const xml = await xmlOf(saved);
		expect(rows(xml)[0]).toContain('w:val="2"');
		expect(rows(xml)[0]).toContain(NOWRAP);
		expect(layoutOf(await reread(saved))).toEqual([['a/b:2'], ['c', 'd']]);
	});
});

describe('saving a restructured table nested in a cell', () => {
	const inner = (rowsXml: string) =>
		`<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>${rowsXml}</w:tbl>`;
	const NESTED =
		`<w:document xmlns:w="${ns}"><w:body><w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLook w:val="04A0"/></w:tblPr>` +
		`<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
		`<w:tr><w:tc><w:tcPr>${tcW(3000)}${NOWRAP}</w:tcPr>${p('outer')}${inner(`<w:tr>${tc('i1', tcW(1500))}${tc('i2', tcW(1500))}</w:tr>`)}<w:p/></w:tc>${tc('B', tcW(3000))}</w:tr>` +
		`<w:tr>${tc('C', tcW(3000))}${tc('D', tcW(3000))}</w:tr></w:tbl><w:sectPr/></w:body></w:document>`;

	it('edits the inner table and leaves the outer table as authored', async () => {
		const { loaded, table } = await open(NESTED);
		const cell = must(table.rows[0]?.[0]);
		const innerTable = must(
			cell.sourceBlocks?.find((block): block is Table => block.type === 'table'),
		);
		expect(innerTable.restructurable).toBe(true);
		const edited = must(insertColumn(innerTable, 0, 'right', newId));
		innerTable.rows = edited.rows;
		innerTable.grid = edited.grid as NonNullable<Table['grid']>;
		innerTable.structureEditable = false;
		const xml = await xmlOf(await loaded.save());
		expect(xml).toContain('<w:tblLook w:val="04A0"/>');
		expect(xml).toContain(NOWRAP);
		const nested = must(xml.match(/<w:tc>(?:(?!<\/w:tc>)[\s\S])*?<w:tbl>[\s\S]*?<\/w:tbl>/))[0];
		expect(nested.match(/<w:gridCol /g)).toHaveLength(3);
		expect(nested.match(/<w:tc>/g)?.length).toBe(4);
		// Outer rows are untouched: same cells, same text, same grid.
		const outerTable = expectTable((await loadDocx(await loaded.save())).model.blocks[0]);
		expect(outerTable.rows.map((row) => row.length)).toEqual([2, 2]);
		expect(outerTable.grid).toEqual([3000, 3000]);
		const savedInner = must(
			outerTable.rows[0]?.[0]?.sourceBlocks?.find(
				(block): block is Table => block.type === 'table',
			),
		);
		expect(layoutOf(savedInner)).toEqual([['i1', '', 'i2']]);
	});
});
