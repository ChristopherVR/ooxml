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
import { cellTarget, drawLineInCell, eraseCellSide, insertDrawnTable } from './table-freeform';
import { tableAt } from './test-support';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const tc = (text: string) =>
	`<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const SIMPLE =
	`<w:document xmlns:w="${ns}"><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
	`<w:tr>${tc('a')}${tc('b')}</w:tr><w:tr>${tc('c')}${tc('d')}</w:tr>` +
	`</w:tbl><w:p><w:r><w:t>after</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;

const views: EditorView[] = [];
afterEach(() => {
	for (const view of views.splice(0)) view.destroy();
	document.body.replaceChildren();
});

async function setup() {
	const zip = new JSZip();
	zip.file('word/document.xml', SIMPLE);
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const doc = modelToDoc(loaded.model);
	const state = EditorState.create({
		doc,
		selection: TextSelection.atStart(doc),
		plugins: [history(), tableStructureGuardPlugin()],
	});
	const host = document.createElement('div');
	document.body.append(host);
	const view = new EditorView(host, { state });
	views.push(view);
	return { loaded, view };
}

/** Position inside the paragraph whose text is `text`. */
function inside(view: EditorView, text: string): number {
	let at = -1;
	view.state.doc.descendants((node, pos) => {
		if (at < 0 && node.type === schema.nodes.paragraph && node.textContent === text) at = pos + 1;
		return at < 0;
	});
	return at;
}

const target = (view: EditorView, text: string) => {
	const found = cellTarget(view.state.doc, inside(view, text));
	if (!found) throw new Error(`No cell around ${text}`);
	return found;
};

function layout(table: Table): string[][] {
	const text = (cell: TableCell) =>
		cell.paragraphs.map((q) => q.runs.map((r) => r.text).join('')).join('/');
	return table.rows.map((row) =>
		row.map(
			(cell) =>
				`${text(cell)}${cell.gridSpan && cell.gridSpan > 1 ? `:${cell.gridSpan}` : ''}${
					cell.verticalMerge === 'restart' ? '^' : cell.verticalMerge === 'continue' ? '|' : ''
				}`,
		),
	);
}

async function saved({ loaded, view }: Awaited<ReturnType<typeof setup>>) {
	const bytes = await loaded.save(docToModel(view.state.doc, loaded.model));
	return tableAt((await loadDocx(bytes)).model.blocks, 0);
}

describe('Draw Table and Eraser in the editor', () => {
	it('splits a cell with a drawn vertical line, saves the new grid and undoes in one step', async () => {
		const ctx = await setup();
		const before = ctx.view.state.doc;
		expect(drawLineInCell(ctx.view, target(ctx.view, 'a'), 'vertical', 0.5)).toBe(true);
		const table = await saved(ctx);
		expect(table.grid).toEqual([1000, 1000, 2000]);
		expect(layout(table)).toEqual([
			['a', '', 'b'],
			['c:2', 'd'],
		]);
		undo(ctx.view.state, ctx.view.dispatch);
		expect(ctx.view.state.doc.eq(before)).toBe(true);
	});

	it('splits a cell into two rows with a drawn horizontal line', async () => {
		const ctx = await setup();
		expect(drawLineInCell(ctx.view, target(ctx.view, 'b'), 'horizontal', 0.5)).toBe(true);
		expect(layout(await saved(ctx))).toEqual([
			['a^', 'b'],
			['|', ''],
			['c', 'd'],
		]);
	});

	it('refuses a line at the edge of a cell', async () => {
		const ctx = await setup();
		const before = ctx.view.state.doc;
		expect(drawLineInCell(ctx.view, target(ctx.view, 'a'), 'vertical', 0.02)).toBe(false);
		expect(ctx.view.state.doc.eq(before)).toBe(true);
	});

	it('merges the cells beside an erased inner line', async () => {
		const ctx = await setup();
		expect(eraseCellSide(ctx.view, target(ctx.view, 'a'), 'right')).toBeUndefined();
		expect(layout(await saved(ctx))).toEqual([['a/b:2'], ['c', 'd']]);
	});

	it('removes the border of an erased outer line', async () => {
		const ctx = await setup();
		expect(eraseCellSide(ctx.view, target(ctx.view, 'a'), 'left')).toBeUndefined();
		const table = await saved(ctx);
		expect(['none', 'nil']).toContain(table.rows[0]?.[0]?.borders?.left?.style);
	});

	it('explains an erase that cannot leave a rectangle', async () => {
		const ctx = await setup();
		expect(eraseCellSide(ctx.view, target(ctx.view, 'c'), 'top')).toBeUndefined();
		// a over c is now one merged cell; b and d beside it do not make a rectangle with it.
		expect(eraseCellSide(ctx.view, target(ctx.view, 'b'), 'left')).toBe('tableTab.eraseRefused');
	});

	it('draws a one-cell table of the dragged size', async () => {
		const ctx = await setup();
		const pos = inside(ctx.view, 'after') - 1;
		expect(insertDrawnTable(ctx.view, pos, 2880, 720)).toBe(true);
		const drawn = docToModel(ctx.view.state.doc, ctx.loaded.model).blocks[1];
		expect(drawn?.type).toBe('table');
		if (drawn?.type !== 'table') return;
		expect(drawn.grid).toEqual([2880]);
		expect(drawn.rowProperties?.[0]).toMatchObject({ heightTwips: 720, heightRule: 'atLeast' });
		expect(insertDrawnTable(ctx.view, pos, 50, 50)).toBe(false);
	});
});
