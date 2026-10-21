import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index';
import { selectListItem } from './content-control-edit';
import type { ContentControl, Table } from './model';
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

	it('writes a cell or row given another control in a control of its own', async () => {
		// A repeating section's item counts as the type element it replaces.
		const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
		const item = await hand(
			`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:sdt><w:sdtPr xmlns:w15="${W15}">` +
				`<w:id w:val="7"/><w15:repeatingSectionItem/></w:sdtPr><w:sdtContent>${row(label('A'))}</w:sdtContent></w:sdt></w:tbl>`,
		);
		const itemRow = tableOf(item.model.blocks).rowProperties![0]!;
		itemRow.contentControl = { ...itemRow.contentControl!, kind: 'text' };
		const itemXml = await documentXml(await item.save(item.model));
		expect(itemXml).toContain('<w:text/>');
		expect(itemXml).not.toContain('repeatingSectionItem');
		// The same control given another kind is patched in place: the properties the model does
		// not hold, such as its data binding, stay.
		const bound = await hand(
			grid(
				row(
					`${label('Status')}${choice('11', '<w:dataBinding w:xpath="/ns0:kind" w:storeItemID="{11111111-2222-3333-4444-555555555555}"/>')}`,
				),
			),
		);
		const kind = tableOf(bound.model.blocks).rows[0]![1]!;
		kind.contentControl = { ...kind.contentControl!, kind: 'comboBox' };
		const boundXml = await documentXml(await bound.save(bound.model));
		expect(boundXml).toContain('w:xpath="/ns0:kind"');
		expect(boundXml).toContain('<w:comboBox>');
		expect(boundXml).not.toContain('dropDownList');
		// Its properties come from the model: a changed title is written with the new id.
		const renamed = await hand(grid(row(`${label('Status')}${choice('11')}`)));
		const own = tableOf(renamed.model.blocks).rows[0]![1]!;
		own.contentControl = { ...own.contentControl!, id: '98', alias: 'Changed' };
		expect(await documentXml(await renamed.save(renamed.model))).toContain(
			'<w:alias w:val="Changed"/><w:id w:val="98"/>',
		);
		// The model can give a cell or row another control, another id or kind, than the w:sdt the
		// file has around it. Patching that w:sdt would keep its old id and put a second type
		// element (w14:checkbox beside w:dropDownList) in its properties.
		const cells = await hand(grid(row(`${label('Status')}${choice('11')}`)));
		const loaded = tableOf(cells.model.blocks).rows[0]![1]!.contentControl!;
		for (const control of [
			{ id: '99', level: 'cell' as const, kind: 'checkbox' as const, checked: true },
			{ ...loaded, kind: 'checkbox' as const, checked: true },
		]) {
			const cell = tableOf(cells.model.blocks).rows[0]![1]!;
			cell.contentControl = control;
			cell.paragraphs[0]!.runs = [{ text: '\u2612' }];
			const xml = await documentXml(await cells.save(cells.model));
			expect(xml).not.toContain('w:dropDownList');
			expect(xml.match(/<w:sdt>/g)).toHaveLength(1);
			expect(xml.match(/<w:id w:val="(\d+)"\/>/g)).toEqual([`<w:id w:val="${control.id}"/>`]);
			const reopened = tableOf((await loadDocx(await cells.save(cells.model))).model.blocks);
			expect(reopened.rows[0]![1]!.contentControl).toMatchObject({
				id: control.id,
				kind: 'checkbox',
				checked: true,
			});
		}
		const line = (content: string) =>
			`<w:sdt><w:sdtPr><w:alias w:val="Item"/><w:id w:val="7"/><w:text/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
		const rows = await hand(grid(line(row(`${label('A')}${label('B')}`))));
		tableOf(rows.model.blocks).rowProperties![0]!.contentControl = {
			id: '98',
			level: 'row',
			kind: 'richText',
			alias: 'Line',
		};
		const rowXml = await documentXml(await rows.save(rows.model));
		expect(rowXml).toContain(
			`</w:tblGrid><w:sdt><w:sdtPr><w:alias w:val="Line"/><w:id w:val="98"/></w:sdtPr><w:sdtContent>${row(`${label('A')}${label('B')}`)}</w:sdtContent></w:sdt></w:tbl>`,
		);
	});

	it('takes a cell given another control out of a control around several cells', async () => {
		// A cell given the control of the cell before it joins its w:sdt, and a new control whose id
		// another w:sdt has, in another row, takes a new one: ids are unique in a document.
		const named = (id: string, text: string) =>
			`<w:sdt><w:sdtPr><w:alias w:val="C${id}"/><w:id w:val="${id}"/></w:sdtPr><w:sdtContent>${label(text)}</w:sdtContent></w:sdt>`;
		const joined = await hand(
			grid(row(`${named('21', 'A')}${named('23', 'B')}`) + row(`${label('C')}${label('D')}`)),
		);
		const rows = tableOf(joined.model.blocks).rows;
		rows[0]![1]!.contentControl = { ...rows[0]![0]!.contentControl! };
		rows[1]![1]!.contentControl = { ...rows[0]![0]!.contentControl! };
		const joinedXml = await documentXml(await joined.save(joined.model));
		expect(joinedXml).toContain(
			`<w:id w:val="21"/></w:sdtPr><w:sdtContent>${label('A')}${label('B')}</w:sdtContent>`,
		);
		expect(joinedXml).not.toContain('<w:id w:val="23"/>');
		const joinedIds = [...joinedXml.matchAll(/<w:id w:val="(\d+)"\/>/g)].map((match) => match[1]);
		expect(joinedIds).toHaveLength(2);
		expect(new Set(joinedIds).size).toBe(2);
		// Cells given together a control whose id another w:sdt has share one, under a new id.
		const both = await hand(
			grid(row(`${named('21', 'A')}${label('B')}`) + row(`${label('C')}${label('D')}`)),
		);
		const bothRows = tableOf(both.model.blocks).rows;
		for (const cell of bothRows[1]!) cell.contentControl = { ...bothRows[0]![0]!.contentControl! };
		const bothXml = await documentXml(await both.save(both.model));
		const bothIds = [...bothXml.matchAll(/<w:id w:val="(\d+)"\/>/g)].map((match) => match[1]);
		expect(bothIds).toHaveLength(2);
		expect(bothIds[0]).toBe('21');
		expect(bothXml).toContain(`</w:sdtPr><w:sdtContent>${label('C')}${label('D')}</w:sdtContent>`);
		// The rest of a split control keeps its new id when its own properties change too.
		const split = await hand(
			`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
				row(
					`<w:sdt><w:sdtPr><w:alias w:val="Group"/><w:id w:val="21"/></w:sdtPr><w:sdtContent>` +
						`${label('A')}${label('B')}${label('C')}</w:sdtContent></w:sdt>`,
				) +
				'</w:tbl>',
		);
		const [left, middle, right] = tableOf(split.model.blocks).rows[0]!;
		middle!.contentControl = { id: '23', level: 'cell', kind: 'richText' };
		for (const cell of [left!, right!])
			cell.contentControl = { ...cell.contentControl!, alias: 'Renamed' };
		const splitIds = [
			...(await documentXml(await split.save(split.model))).matchAll(/<w:id w:val="(\d+)"\/>/g),
		].map((match) => match[1]);
		expect(splitIds).toHaveLength(3);
		expect(new Set(splitIds).size).toBe(3);
		// A cell after the rest of a split control given that control joins the rest.
		const after = await hand(
			`<w:tbl><w:tblPr/><w:tblGrid>${'<w:gridCol w:w="1500"/>'.repeat(4)}</w:tblGrid>` +
				row(
					`<w:sdt><w:sdtPr><w:alias w:val="Group"/><w:id w:val="21"/></w:sdtPr><w:sdtContent>` +
						`${label('A')}${label('B')}${label('C')}</w:sdtContent></w:sdt>${label('D')}`,
				) +
				'</w:tbl>',
		);
		const afterCells = tableOf(after.model.blocks).rows[0]!;
		afterCells[1]!.contentControl = { id: '23', level: 'cell', kind: 'richText' };
		afterCells[3]!.contentControl = { ...afterCells[0]!.contentControl! };
		const afterXml = await documentXml(await after.save(after.model));
		expect([...afterXml.matchAll(/<w:id w:val="(\d+)"\/>/g)]).toHaveLength(3);
		expect(afterXml).toContain(`</w:sdtPr><w:sdtContent>${label('C')}${label('D')}</w:sdtContent>`);
		// Split twice, the rest keeps the control it is the rest of: a cell after it given that joins.
		const twice = await hand(
			`<w:tbl><w:tblPr/><w:tblGrid>${'<w:gridCol w:w="1000"/>'.repeat(6)}</w:tblGrid>` +
				row(
					`<w:sdt><w:sdtPr><w:alias w:val="Group"/><w:id w:val="21"/></w:sdtPr><w:sdtContent>` +
						`${['A', 'B', 'C', 'D', 'E'].map(label).join('')}</w:sdtContent></w:sdt>${label('F')}`,
				) +
				'</w:tbl>',
		);
		const twiceCells = tableOf(twice.model.blocks).rows[0]!;
		twiceCells[1]!.contentControl = { id: '23', level: 'cell', kind: 'richText' };
		twiceCells[3]!.contentControl = { id: '25', level: 'cell', kind: 'richText' };
		twiceCells[5]!.contentControl = { ...twiceCells[0]!.contentControl! };
		const twiceXml = await documentXml(await twice.save(twice.model));
		expect(new Set([...twiceXml.matchAll(/<w:id w:val="(\d+)"\/>/g)].map((m) => m[1])).size).toBe(
			5,
		);
		expect(twiceXml).toContain(`</w:sdtPr><w:sdtContent>${label('E')}${label('F')}</w:sdtContent>`);
		// Cells given the same new control together keep one control, as they shared the old one.
		const pair = await hand(
			grid(
				row(
					`<w:sdt><w:sdtPr><w:alias w:val="Group"/><w:id w:val="21"/></w:sdtPr><w:sdtContent>` +
						`${label('A')}${label('B')}</w:sdtContent></w:sdt>`,
				),
			),
		);
		for (const cell of tableOf(pair.model.blocks).rows[0]!)
			cell.contentControl = { ...cell.contentControl!, id: '99', alias: 'Moved' };
		const pairXml = await documentXml(await pair.save(pair.model));
		expect(pairXml.match(/<w:sdt>/g)).toHaveLength(1);
		expect(pairXml).toContain(
			`<w:id w:val="99"/></w:sdtPr><w:sdtContent>${label('A')}${label('B')}`,
		);
		expect(pairXml).toContain('<w:alias w:val="Moved"/>');
		const wrap = (content: string) =>
			`<w:sdt><w:sdtPr><w:alias w:val="Group"/><w:id w:val="21"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
		// At the start or end it moves out beside the control; from its middle, the cells after it
		// go to a copy under a new id, since ids are unique (a repeated one would not survive Word).
		for (const [at, expected] of [
			[0, ['23', '21', '21']],
			[1, ['21', '23', 'new']],
			[2, ['21', '21', '23']],
		] as const) {
			const loaded = await hand(
				`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
					`${row(wrap(`${label('A')}${label('B')}${label('C')}`))}</w:tbl>`,
			);
			tableOf(loaded.model.blocks).rows[0]![at]!.contentControl = {
				id: '23',
				level: 'cell',
				kind: 'richText',
				alias: 'Own',
			};
			const xml = await documentXml(await loaded.save(loaded.model));
			expect(xml).toContain(`${label('A')}`);
			const reopened = tableOf((await loadDocx(await loaded.save(loaded.model))).model.blocks);
			const cells = reopened.rows[0]!;
			expect(cells.map((cell) => cell.paragraphs[0]!.runs[0]!.text)).toEqual(['A', 'B', 'C']);
			const ids = cells.map((cell) => cell.contentControl!.id);
			if (at !== 1) expect(ids).toEqual(expected);
			else {
				expect(ids.slice(0, 2)).toEqual(['21', '23']);
				expect(ids[2]).toMatch(/^\d+$/);
				expect(['21', '23']).not.toContain(ids[2]);
			}
			expect(cells.map((cell) => cell.contentControl?.alias)).toEqual(
				expected.map((id) => (id === '23' ? 'Own' : 'Group')),
			);
		}
	});

	it('takes a row or cell given another control out of one that holds range markers beside it', async () => {
		// A row or cell control's content can hold range markers beside its rows or cells
		// (EG_ContentRowContent, EG_ContentCellContent). They are no rows or cells: what is taken out
		// takes those on its side along, so no control is left holding only markers.
		const old = (content: string) =>
			`<w:sdt><w:sdtPr><w:alias w:val="Old"/><w:id w:val="31"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
		const start = '<w:bookmarkStart w:id="8" w:name="Kept"/>';
		const end = '<w:bookmarkEnd w:id="8"/>';
		const fresh = '<w:sdt><w:sdtPr><w:id w:val="33"/></w:sdtPr><w:sdtContent>';
		const given = async (body: string, r: number, c: number) => {
			const loaded = await hand(body);
			tableOf(loaded.model.blocks).rows[r]![c]!.contentControl = {
				id: '33',
				level: 'cell',
				kind: 'richText',
			};
			return documentXml(await loaded.save(loaded.model));
		};
		// The only cell of the control, between the markers: the control goes, they stay around it.
		const only = await given(
			grid(row(`${label('Status')}${old(start + label('Open') + end)}`)),
			0,
			1,
		);
		expect(only.match(/<w:sdt>/g)).toHaveLength(1);
		expect(only).toContain(
			`${label('Status')}${start}${fresh}${label('Open')}</w:sdtContent></w:sdt>${end}</w:tr>`,
		);
		// The first cell goes before it with the marker before it, the last after it with the one after.
		const both = grid(row(old(start + label('A') + label('B') + end)));
		const first = await given(both, 0, 0);
		expect(first.match(/<w:sdt>/g)).toHaveLength(2);
		expect(first).toContain(
			`<w:tr>${start}${fresh}${label('A')}</w:sdtContent></w:sdt><w:sdt><w:sdtPr><w:alias w:val="Old"/>` +
				`<w:id w:val="31"/></w:sdtPr><w:sdtContent>${label('B')}${end}</w:sdtContent></w:sdt></w:tr>`,
		);
		const last = await given(both, 0, 1);
		expect(last.match(/<w:sdt>/g)).toHaveLength(2);
		expect(last).toContain(
			`<w:tr><w:sdt><w:sdtPr><w:alias w:val="Old"/><w:id w:val="31"/></w:sdtPr><w:sdtContent>` +
				`${start}${label('A')}</w:sdtContent></w:sdt>${fresh}${label('B')}</w:sdtContent></w:sdt>${end}</w:tr>`,
		);
		// A cell given the control of the cell before it joins it, with the markers between them.
		const joined = await hand(
			grid(
				row(
					`<w:sdt><w:sdtPr><w:alias w:val="C21"/><w:id w:val="21"/></w:sdtPr><w:sdtContent>${label('A')}` +
						`</w:sdtContent></w:sdt>${old(start + label('B') + end)}`,
				),
			),
		);
		const [left, right] = tableOf(joined.model.blocks).rows[0]!;
		right!.contentControl = { ...left!.contentControl! };
		const joinedXml = await documentXml(await joined.save(joined.model));
		expect(joinedXml).toContain(
			`<w:tr><w:sdt><w:sdtPr><w:alias w:val="C21"/><w:id w:val="21"/></w:sdtPr><w:sdtContent>` +
				`${label('A')}${start}${label('B')}</w:sdtContent></w:sdt>${end}</w:tr>`,
		);
		// A row the same way.
		const rows = await hand(grid(old(start + row(`${label('A')}${label('B')}`) + end)));
		tableOf(rows.model.blocks).rowProperties![0]!.contentControl = {
			id: '33',
			level: 'row',
			kind: 'richText',
		};
		const rowXml = await documentXml(await rows.save(rows.model));
		expect(rowXml.match(/<w:sdt>/g)).toHaveLength(1);
		expect(rowXml).toContain(
			`</w:tblGrid>${start}${fresh}${row(`${label('A')}${label('B')}`)}</w:sdtContent></w:sdt>${end}</w:tbl>`,
		);
	});

	it('writes only valid properties for a control the model gives a cell', async () => {
		// A cell pasted from another page, a remote peer or the model API can hold any values: what
		// Word cannot read is not written (w:id is optional, so a bad id is left out).
		const controls = [
			{ id: 'abc', kind: 'bogus', lock: 'everything' },
			{ id: '99999999999', kind: 'unknown' },
			{ id: 'sdt-gen-5', kind: 'repeatingSection' },
			{
				id: '12',
				kind: 'checkbox',
				checkedGlyph: { code: 'zz' },
				uncheckedGlyph: { code: '2610' },
			},
			{ id: '13', kind: 'dropDownList', items: 5 },
			{ id: '14', kind: 'date', fullDate: 'soon', dateFormat: 'd/M/yyyy' },
		];
		const loaded = await hand(
			grid(controls.map((_, index) => row(`${label(`L${index}`)}${label(`V${index}`)}`)).join('')),
		);
		const table = tableOf(loaded.model.blocks);
		controls.forEach((control, index) => {
			table.rows[index]![1]!.contentControl = {
				level: 'cell',
				...control,
			} as unknown as ContentControl;
		});
		const xml = await documentXml(await loaded.save(loaded.model));
		const properties = [...xml.matchAll(/<w:sdtPr(?:\/>|>(.*?)<\/w:sdtPr>)/g)].map(
			(match) => match[1] ?? '',
		);
		expect(properties).toHaveLength(controls.length);
		const ids = properties.map((sdtPr) => /<w:id w:val="([^"]*)"\/>/.exec(sdtPr)?.[1]);
		expect(ids).toEqual([undefined, undefined, undefined, '12', '13', '14']);
		expect(xml).not.toMatch(
			/w:bogus|w:unknown|w:repeatingSection|w:lock|"zz"|w:fullDate|w:listItem/,
		);
		expect(properties[3]).toContain('<w14:uncheckedState w14:val="2610"/>');
		expect(properties[5]).toContain('<w:date><w:dateFormat w:val="d/M/yyyy"/></w:date>');
		const reopened = tableOf((await loadDocx(await loaded.save(loaded.model))).model.blocks);
		expect(reopened.rows.map((cells) => cells[1]!.contentControl?.kind)).toEqual([
			'richText',
			'richText',
			'richText',
			'checkbox',
			'dropDownList',
			'date',
		]);
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
