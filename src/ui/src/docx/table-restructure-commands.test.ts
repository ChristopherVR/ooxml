// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, type Table, type TableCell } from 'ooxml-core/docx';
import { tableStructureGuardPlugin } from 'ooxml-core/docx/ui';
import { history, undo } from 'prosemirror-history';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { docToModel, modelToDoc } from './model-adapter';
import { schema } from './schema';
import { canExecuteTableCommand, executeTableCommand, type TableCommand } from './table-commands';
import { tableAt } from './test-support';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const tc = (text: string, props = '') =>
	`<w:tc><w:tcPr>${props}</w:tcPr>${text ? p(text) : '<w:p/>'}</w:tc>`;
const w = (n: number) => `<w:tcW w:w="${n}" w:type="dxa"/>`;
const grid = (cols: number[]) =>
	`<w:tblGrid>${cols.map((c) => `<w:gridCol w:w="${c}"/>`).join('')}</w:tblGrid>`;

/** A | B (A spans two columns) over C | D | E, with E merged down over F. */
const MERGED =
	`<w:document xmlns:w="${ns}"><w:body><w:tbl><w:tblPr/>${grid([2000, 2000, 2000])}` +
	`<w:tr>${tc('A', w(4000) + '<w:gridSpan w:val="2"/><w:noWrap/>')}${tc('B', w(2000) + '<w:vMerge w:val="restart"/>')}</w:tr>` +
	`<w:tr>${tc('C', w(2000))}${tc('D', w(2000))}${tc('', w(2000) + '<w:vMerge/>')}</w:tr>` +
	`</w:tbl><w:sectPr/></w:body></w:document>`;
const SIMPLE =
	`<w:document xmlns:w="${ns}"><w:body><w:tbl><w:tblPr/>${grid([2000, 2000])}` +
	`<w:tr>${tc('a', w(2000))}${tc('b', w(2000))}</w:tr><w:tr>${tc('c', w(2000))}${tc('d', w(2000))}</w:tr>` +
	`</w:tbl><w:sectPr/></w:body></w:document>`;
const INNER_OUTER = (inner: string) =>
	`<w:document xmlns:w="${ns}"><w:body><w:tbl><w:tblPr/>${grid([4000, 4000])}<w:tr>` +
	`<w:tc><w:tcPr>${w(4000)}</w:tcPr>${p('outer')}${inner}<w:p/></w:tc>${tc('B', w(4000))}</w:tr></w:tbl><w:sectPr/></w:body></w:document>`;

const views: EditorView[] = [];
afterEach(() => {
	for (const view of views.splice(0)) view.destroy();
	document.body.replaceChildren();
});

async function setup(xml: string, selectText: string) {
	const zip = new JSZip();
	zip.file('word/document.xml', xml);
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const doc = modelToDoc(loaded.model);
	let at = -1;
	doc.descendants((node, pos) => {
		if (at < 0 && node.type === schema.nodes.paragraph && node.textContent === selectText)
			at = pos + 1;
		return at < 0;
	});
	const state = EditorState.create({
		doc,
		selection: TextSelection.create(doc, at),
		plugins: [history(), tableStructureGuardPlugin()],
	});
	const host = document.createElement('div');
	document.body.append(host);
	const view = new EditorView(host, { state });
	views.push(view);
	return { loaded, view };
}

/** The tables of the editor document laid out as text: `A:2`, `B^`, `|` for a continuation. */
function layout(table: Table): string[][] {
	const text = (cell: TableCell) =>
		cell.paragraphs
			.map((q) => q.runs.map((r) => r.text).join(''))
			.filter(Boolean)
			.join('/');
	return table.rows.map((row) =>
		row.map(
			(cell) =>
				`${text(cell)}${cell.gridSpan && cell.gridSpan > 1 ? `:${cell.gridSpan}` : ''}${
					cell.verticalMerge === 'restart' ? '^' : cell.verticalMerge === 'continue' ? '|' : ''
				}`,
		),
	);
}

async function roundTrip({ loaded, view }: Awaited<ReturnType<typeof setup>>) {
	const model = docToModel(view.state.doc, loaded.model);
	const bytes = await loaded.save(model);
	const xml =
		(await (await JSZip.loadAsync(bytes)).file('word/document.xml')?.async('string')) ?? '';
	return { saved: tableAt((await loadDocx(bytes)).model.blocks, 0), xml, model };
}

describe('table commands on merged tables', () => {
	it('enables the row and column commands that used to be blocked', async () => {
		const { view } = await setup(MERGED, 'C');
		for (const action of ['rowBefore', 'rowAfter', 'deleteRow', 'columnAfter'] as TableCommand[])
			expect(canExecuteTableCommand(view, action), action).toBe(true);
	});

	it('inserts a row through the vertical merge and saves it', async () => {
		const ctx = await setup(MERGED, 'C');
		expect(executeTableCommand(ctx.view, 'rowBefore')).toBe(true);
		const { saved, xml } = await roundTrip(ctx);
		expect(layout(saved)).toEqual([
			['A:2', 'B^'],
			['', '', '|'],
			['C', 'D', '|'],
		]);
		expect(xml).toContain('<w:noWrap/>');
	});

	it('does not extend a vertical merge when a row is added below its last row', async () => {
		const ctx = await setup(MERGED, 'C');
		expect(executeTableCommand(ctx.view, 'rowAfter')).toBe(true);
		expect(layout((await roundTrip(ctx)).saved)[2]).toEqual(['', '', '']);
	});

	it('undoes a structure command in one step', async () => {
		const ctx = await setup(MERGED, 'C');
		const before = ctx.view.state.doc;
		executeTableCommand(ctx.view, 'rowBefore');
		expect(ctx.view.state.doc.eq(before)).toBe(false);
		undo(ctx.view.state, ctx.view.dispatch);
		expect(ctx.view.state.doc.eq(before)).toBe(true);
	});

	it('deletes a column and a row', async () => {
		// Word cannot delete a column through a merged cell as a plain grid: it is refused.
		const refused = await setup(MERGED, 'C');
		expect(canExecuteTableCommand(refused.view, 'deleteColumn')).toBe(false);
		const column = await setup(MERGED, 'B');
		expect(executeTableCommand(column.view, 'deleteColumn')).toBe(true);
		expect(layout((await roundTrip(column)).saved)).toEqual([['A:2'], ['C', 'D']]);
		const row = await setup(MERGED, 'C');
		expect(executeTableCommand(row.view, 'deleteRow')).toBe(true);
		expect(layout((await roundTrip(row)).saved)).toEqual([['A:2', 'B']]);
	});

	it('merges the cells a selection spans', async () => {
		const ctx = await setup(MERGED, 'C');
		const { doc } = ctx.view.state;
		let from = -1;
		let to = -1;
		doc.descendants((node, pos) => {
			if (node.type !== schema.nodes.paragraph) return true;
			if (node.textContent === 'C') from = pos + 1;
			if (node.textContent === 'D') to = pos + 1;
			return false;
		});
		ctx.view.dispatch(ctx.view.state.tr.setSelection(TextSelection.create(doc, from, to)));
		expect(executeTableCommand(ctx.view, 'mergeCells')).toBe(true);
		expect(layout((await roundTrip(ctx)).saved)).toEqual([
			['A:2', 'B^'],
			['C/D:2', '|'],
		]);
	});

	it('splits a cell into columns', async () => {
		const ctx = await setup(MERGED, 'A');
		expect(executeTableCommand(ctx.view, 'splitCell', { rows: 1, columns: 2 })).toBe(true);
		expect(layout((await roundTrip(ctx)).saved)[0]).toEqual(['A', '', 'B^']);
	});

	it('guards the merged table against edits that are not structure commands', async () => {
		const ctx = await setup(MERGED, 'C');
		const before = ctx.view.state.doc;
		let from = -1;
		before.descendants((node, pos) => {
			if (from < 0 && node.type === schema.nodes.tableCell) from = pos;
			return from < 0;
		});
		const cell = before.nodeAt(from)!;
		ctx.view.dispatch(ctx.view.state.tr.delete(from, from + cell.nodeSize));
		expect(ctx.view.state.doc.eq(before)).toBe(true);
	});
});

describe('merge and split on a plain table', () => {
	it('merges two cells and saves the merged table', async () => {
		const ctx = await setup(SIMPLE, 'a');
		const { doc } = ctx.view.state;
		let to = -1;
		doc.descendants((node, pos) => {
			if (node.type === schema.nodes.paragraph && node.textContent === 'b') to = pos + 1;
			return to < 0;
		});
		ctx.view.dispatch(
			ctx.view.state.tr.setSelection(TextSelection.create(doc, ctx.view.state.selection.from, to)),
		);
		expect(canExecuteTableCommand(ctx.view, 'mergeCells')).toBe(true);
		expect(executeTableCommand(ctx.view, 'mergeCells')).toBe(true);
		expect(ctx.view.state.doc.child(0).attrs.structureEditable).toBe(false);
		const { saved } = await roundTrip(ctx);
		expect(layout(saved)).toEqual([['a/b:2'], ['c', 'd']]);
	});

	it('keeps plain tables on the rectangular row and column commands', async () => {
		const ctx = await setup(SIMPLE, 'a');
		expect(executeTableCommand(ctx.view, 'columnAfter')).toBe(true);
		expect(ctx.view.state.doc.child(0).attrs.structureEditable).toBe(true);
	});
});

describe('table commands on a nested table', () => {
	const INNER = `<w:tbl><w:tblPr/>${grid([2000, 2000])}<w:tr>${tc('i1', w(2000))}${tc('i2', w(2000))}</w:tr></w:tbl>`;

	it('edits the inner table and leaves the outer table untouched', async () => {
		const ctx = await setup(INNER_OUTER(INNER), 'i1');
		expect(executeTableCommand(ctx.view, 'columnAfter')).toBe(true);
		expect(executeTableCommand(ctx.view, 'rowAfter')).toBe(true);
		const { saved } = await roundTrip(ctx);
		expect(saved.rows.map((row) => row.length)).toEqual([2]);
		expect(saved.grid).toEqual([4000, 4000]);
		const inner = saved.rows[0]![0]!.sourceBlocks!.find(
			(block): block is Table => block.type === 'table',
		)!;
		expect(layout(inner)).toEqual([
			['i1', '', 'i2'],
			['', '', ''],
		]);
		expect(inner.grid).toEqual([2000, 2000, 2000]);
	});

	it('merges cells of the inner table', async () => {
		const ctx = await setup(INNER_OUTER(INNER), 'i1');
		const { doc } = ctx.view.state;
		let to = -1;
		doc.descendants((node, pos) => {
			if (node.type === schema.nodes.paragraph && node.textContent === 'i2') to = pos + 1;
			return to < 0;
		});
		ctx.view.dispatch(
			ctx.view.state.tr.setSelection(TextSelection.create(doc, ctx.view.state.selection.from, to)),
		);
		expect(executeTableCommand(ctx.view, 'mergeCells')).toBe(true);
		const { saved } = await roundTrip(ctx);
		const inner = saved.rows[0]![0]!.sourceBlocks!.find(
			(block): block is Table => block.type === 'table',
		)!;
		expect(layout(inner)).toEqual([['i1/i2:2']]);
		expect(layout(saved)[0]![1]).toBe('B');
	});
});

describe('a row covered entirely by vertical merges', () => {
	const COVERED =
		`<w:document xmlns:w="${ns}"><w:body><w:tbl><w:tblPr/>${grid([2000, 2000])}` +
		`<w:tr>${tc('a', w(2000) + '<w:vMerge w:val="restart"/>')}${tc('b', w(2000) + '<w:vMerge w:val="restart"/>')}</w:tr>` +
		`<w:tr>${tc('', w(2000) + '<w:vMerge/>')}${tc('', w(2000) + '<w:vMerge/>')}</w:tr>` +
		`</w:tbl><w:sectPr/></w:body></w:document>`;

	it('opens as a row without cells and saves the merge again', async () => {
		const ctx = await setup(COVERED, 'a');
		const table = ctx.view.state.doc.firstChild!;
		expect(table.child(1).childCount).toBe(0);
		const { saved } = await roundTrip(ctx);
		expect(layout(saved)).toEqual([
			['a^', 'b^'],
			['|', '|'],
		]);
	});
});
