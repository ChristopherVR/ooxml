// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { schema } from './schema';
import { cellTab } from './table-navigation';

const cell = (id: string, text: string) =>
	schema.nodes.tableCell!.create(null, schema.nodes.paragraph!.create({ id }, schema.text(text)));

function setup() {
	const table = schema.nodes.table!.create({ id: 't' }, [
		schema.nodes.tableRow!.create(null, [cell('a', 'A1'), cell('b', 'B1')]),
		schema.nodes.tableRow!.create(null, [cell('c', 'A2'), cell('d', 'B2')]),
	]);
	const doc = schema.nodes.doc!.create(null, [table, schema.nodes.paragraph!.create({ id: 'z' })]);
	const view = new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, schema }),
	});
	return view;
}
const cellText = (view: EditorView) =>
	view.state.doc.textBetween(view.state.selection.from, view.state.selection.to);

it('moves through cells in reading order and selects their text', () => {
	const view = setup();
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 3)));
	expect(cellText(view)).toBe('');
	cellTab(1)(view.state, view.dispatch, view);
	expect(cellText(view)).toBe('B1');
	cellTab(1)(view.state, view.dispatch, view);
	expect(cellText(view)).toBe('A2');
	cellTab(-1)(view.state, view.dispatch, view);
	expect(cellText(view)).toBe('B1');
	view.destroy();
});

it('adds a row when Tab leaves the last cell and ignores text outside tables', () => {
	const view = setup();
	let pos = 0;
	view.state.doc.descendants((node, at) => {
		if (node.isText && node.text === 'B2') pos = at + 1;
	});
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, pos)));
	expect(cellTab(1)(view.state, view.dispatch, view)).toBe(true);
	expect(view.state.doc.child(0).childCount).toBe(3);
	view.dispatch(view.state.tr.setSelection(TextSelection.atEnd(view.state.doc)));
	expect(cellTab(1)(view.state, view.dispatch, view)).toBe(false);
	view.destroy();
});
