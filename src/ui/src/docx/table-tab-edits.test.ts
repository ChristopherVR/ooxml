// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { createDocument, loadDocx, saveDocx, type Table } from 'ooxml-core/docx';
import { autoFitTable } from 'ooxml-core/docx/ui';
import { history, undo } from 'prosemirror-history';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { docToModel, modelToDoc } from './model-adapter';
import { tableCellSelectionPlugin } from './table-cell-selection';
import { paintCellBorder } from './table-border-commands';
import { tableTabContext } from './table-tab-context';
import { alignTableCells, patchTableNode, selectTableParts } from './table-tab-edits';
import { runRibbonCommand } from './editor-commands';

let view: EditorView | undefined;
afterEach(() => {
	view?.destroy();
	view = undefined;
	document.body.replaceChildren();
});

function setup() {
	const model = createDocument();
	const table: Table = {
		type: 'table',
		id: 'table',
		rows: [
			['a', 'b'],
			['c', 'd'],
		].map((row) =>
			row.map((text) => ({ paragraphs: [{ type: 'paragraph', id: text, runs: [{ text }] }] })),
		),
	};
	model.blocks = [table];
	const doc = modelToDoc(model);
	const editor = new EditorView(document.body.appendChild(document.createElement('div')), {
		handleScrollToSelection: () => true,
		state: EditorState.create({
			doc,
			selection: TextSelection.create(doc, 4),
			plugins: [history(), tableCellSelectionPlugin()],
		}),
	});
	view = editor;
	return { editor, model, table, doc };
}

it('aligns the two cells of a selected column as a single undo step', async () => {
	const { editor, doc, model } = setup();
	selectTableParts(editor, tableTabContext(editor.state)!, 'column');
	expect(editor.state.selection.ranges).toHaveLength(2);
	expect(editor.dom.querySelectorAll('td.dve-selected-table-cell')).toHaveLength(2);
	alignTableCells(editor, tableTabContext(editor.state)!, {
		type: 'tableTab',
		key: 'align',
		vertical: 'bottom',
		horizontal: 'right',
	});
	const result = docToModel(editor.state.doc, model, doc);
	const table = (await loadDocx(await saveDocx(result))).model.blocks[0];
	expect(table?.type).toBe('table');
	if (table?.type !== 'table') return;
	expect(table.rows.flat().map((cell) => cell.verticalAlign)).toEqual([
		'bottom',
		undefined,
		'bottom',
		undefined,
	]);
	expect(table.rows.flat().map((cell) => cell.paragraphs[0]?.align)).toEqual([
		'right',
		undefined,
		'right',
		undefined,
	]);
	expect(undo(editor.state, editor.dispatch)).toBe(true);
	expect(editor.state.doc.eq(doc)).toBe(true);
});

it('paints the shared edge, applies AutoFit and keeps both changes after reopening', async () => {
	const { editor, doc, model, table } = setup();
	const context = tableTabContext(editor.state)!;
	paintCellBorder(editor, context.pos, 0, 0, 'right', {
		style: 'double',
		sizeEighthPoints: 12,
		color: '#4472C4',
	});
	const painted = docToModel(editor.state.doc, model, doc);
	const current = painted.blocks[0];
	expect(current?.type).toBe('table');
	if (current?.type !== 'table') return;
	const active = tableTabContext(editor.state)!;
	editor.dispatch(
		patchTableNode(editor.state.tr, active.pos, active.node, autoFitTable(current, 'window')),
	);
	const next = docToModel(editor.state.doc, model);
	const saved = (await loadDocx(await saveDocx(next))).model.blocks[0];
	if (saved?.type !== 'table') throw new Error('Expected a table');
	expect(saved.layout).toBe('autofit');
	expect(saved.widthPercent).toBe(1);
	expect(saved.rows[0]?.[0]?.borders?.right).toEqual(saved.rows[0]?.[1]?.borders?.left);
	expect(saved.rows[0]?.[0]?.borders?.right?.style).toBe('double');
	expect(table.rows[0]?.[0]?.borders).toBeUndefined();
});

it('formats and copies every selected cell without including neighboring cells', () => {
	const { editor, model } = setup();
	selectTableParts(editor, tableTabContext(editor.state)!, 'column');
	runRibbonCommand(editor, { type: 'format', key: 'bold' });
	runRibbonCommand(editor, { type: 'font', key: 'size', value: '22' });
	runRibbonCommand(editor, { type: 'align', value: 'right' });
	const table = docToModel(editor.state.doc, model).blocks[0];
	if (table?.type !== 'table') throw new Error('Expected a table');
	const paragraphs = table.rows.flatMap((row) => row.flatMap((cell) => cell.paragraphs));
	expect(paragraphs.map((p) => [p.align, p.runs[0]?.bold, p.runs[0]?.fontSize])).toEqual([
		['right', true, 22],
		[undefined, undefined, undefined],
		['right', true, 22],
		[undefined, undefined, undefined],
	]);
	const clipboard = editor.serializeForClipboard(editor.state.selection.content());
	expect(clipboard.text.trim().split(/\s+/)).toEqual(['a', 'c']);
	expect([...clipboard.dom.querySelectorAll('td')].map((cell) => cell.textContent)).toEqual([
		'a',
		'c',
	]);
});
