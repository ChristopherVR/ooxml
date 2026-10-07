import { expect, it } from 'vitest';
import { Schema, type Node } from 'prosemirror-model';
import { EditorState, type Transaction, TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { history, undo, redo } from 'prosemirror-history';
import { markSpecs } from './schema-marks';
import {
	acceptAllChanges,
	rejectAllChanges,
	acceptChangeAtCursor,
	collectRevisionRanges,
	rejectRevisionRange,
} from './review-commands';

const schema = new Schema({
	nodes: {
		doc: { content: 'block+', attrs: { sections: { default: null } } },
		paragraph: {
			content: 'text*',
			group: 'block',
			attrs: { id: { default: '' }, markRevision: { default: null }, align: { default: null } },
		},
		table: { content: 'paragraph+', group: 'block' },
		text: {},
	},
	marks: markSpecs,
});
function paragraph(id: string, kind?: 'insert' | 'delete', author = 'Ada'): Node {
	return schema.node(
		'paragraph',
		{
			id,
			align: id === 'last' ? 'right' : 'left',
			...(kind ? { markRevision: { id: 'mark', kind, author } } : {}),
		},
		schema.text(id, [schema.marks.bold!.create()]),
	);
}
function view(nodes: Node[], sections: string | null = null): EditorView {
	const editor = {
		state: EditorState.create({
			doc: schema.node('doc', { sections }, nodes),
			plugins: [history()],
		}),
		editable: true,
		dispatch(tr: Transaction) {
			this.state = this.state.apply(tr);
		},
		focus() {},
	};
	return editor as unknown as EditorView;
}
for (const kind of ['insert', 'delete'] as const)
	for (const mode of ['accept', 'reject'] as const)
		it(`${mode}s a ${kind} paragraph boundary and restores it with undo`, () => {
			const editor = view([paragraph('first', kind), paragraph('last')]);
			const initial = editor.state.doc;
			expect(collectRevisionRanges(initial)).toHaveLength(1);
			expect((mode === 'accept' ? acceptAllChanges : rejectAllChanges)(editor)).toBe(true);
			const keep = (mode === 'accept') === (kind === 'insert');
			expect(editor.state.doc.childCount).toBe(keep ? 2 : 1);
			expect(editor.state.doc.textContent).toBe('firstlast');
			expect(collectRevisionRanges(editor.state.doc)).toEqual([]);
			if (!keep) {
				expect(editor.state.doc.firstChild!.attrs.id).toBe('last');
				expect(editor.state.doc.firstChild!.attrs.align).toBe('right');
				expect(editor.state.doc.firstChild!.firstChild!.marks[0]!.type.name).toBe('bold');
			}
			const resolved = editor.state.doc;
			expect(undo(editor.state, editor.dispatch.bind(editor))).toBe(true);
			expect(editor.state.doc.eq(initial)).toBe(true);
			expect(redo(editor.state, editor.dispatch.bind(editor))).toBe(true);
			expect(editor.state.doc.eq(resolved)).toBe(true);
		});
it('merges consecutive inserted boundaries from the end with final paragraph formatting', () => {
	const editor = view([
		paragraph('first', 'insert'),
		paragraph('second', 'insert'),
		paragraph('last'),
	]);
	expect(rejectAllChanges(editor)).toBe(true);
	expect(editor.state.doc.childCount).toBe(1);
	expect(editor.state.doc.firstChild!.attrs).toMatchObject({ id: 'last', align: 'right' });
	expect(editor.state.doc.textContent).toBe('firstsecondlast');
});
it('resolves and navigates a paragraph-only change at the end of its text', () => {
	const editor = view([paragraph('first', 'insert'), paragraph('last')]);
	const range = collectRevisionRanges(editor.state.doc)[0]!;
	editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, range.from)));
	expect(acceptChangeAtCursor(editor)).toBe(true);
	expect(editor.state.doc.childCount).toBe(2);
	expect(collectRevisionRanges(editor.state.doc)).toEqual([]);
});
it('keeps colliding paragraph revision ids from different authors separate', () => {
	const editor = view([
		paragraph('first', 'insert'),
		paragraph('second', 'insert', 'Bob'),
		paragraph('last'),
	]);
	rejectRevisionRange(editor, collectRevisionRanges(editor.state.doc)[0]!);
	expect(editor.state.doc.childCount).toBe(2);
	expect(collectRevisionRanges(editor.state.doc).map((range) => range.author)).toEqual(['Bob']);
});
for (const boundary of ['last', 'table', 'section'])
	it(`preserves the document when removing an unsupported ${boundary} boundary`, () => {
		const nodes = [paragraph('first', 'insert')];
		if (boundary === 'table') nodes.push(schema.node('table', null, paragraph('last')));
		if (boundary === 'section') nodes.push(paragraph('last'));
		const editor = view(nodes, boundary === 'section' ? '[{"endsAtBlockId":"first"}]' : null);
		const initial = editor.state.doc;
		expect(() => rejectAllChanges(editor)).toThrow('Cannot resolve');
		expect(editor.state.doc).toBe(initial);
		expect(undo(editor.state, editor.dispatch.bind(editor))).toBe(false);
	});
