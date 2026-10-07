import { expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { fieldMarkerNodeSpec } from './break-note-schema';
import { markSpecs } from './schema-marks';
import { runToInlineNodes } from './run-adapter';
import { commentIdsFromNode } from './comment-anchors';
import { addComment } from './comment-commands';
import { commentSelectionRange } from './comment-selection';
const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*' },
		text: { group: 'inline' },
		fieldMarker: fieldMarkerNodeSpec,
	},
	marks: markSpecs,
});
const doc = schema.node(
	'doc',
	null,
	schema.node('paragraph', null, [
		schema.text('L'),
		...runToInlineNodes({ text: '', fieldChar: 'begin' }, schema),
		...runToInlineNodes({ text: '', fieldCode: ' QUOTE ABCDE ' }, schema),
		...runToInlineNodes({ text: '', fieldChar: 'separate' }, schema),
		schema.text('ABCDE'),
		...runToInlineNodes({ text: '', fieldChar: 'end' }, schema),
		schema.text('R'),
	]),
);
for (const [from, to] of [
	[2, 3],
	[6, 7],
	[9, 10],
])
	it(`anchors selection ${from}:${to} to the complete complex field`, () => {
		expect(commentSelectionRange(doc, from!, to!)).toEqual({ from: 2, to: 11 });
		const host = {
			state: EditorState.create({ doc, selection: TextSelection.create(doc, from!, to!) }),
			editable: true,
			dispatch(tr: Transaction) {
				this.state = this.state.apply(tr);
			},
		};
		const view = host as unknown as EditorView;
		addComment(view, 'Ada', 'A', () => 'a');
		addComment(view, 'Bob', 'B', () => 'b');
		const anchors: Array<{ pos: number; ids: string[] }> = [];
		host.state.doc.descendants((node, pos) => {
			if (node.isInline) anchors.push({ pos, ids: commentIdsFromNode(node) });
		});
		expect(anchors).toEqual([
			{ pos: 1, ids: [] },
			{ pos: 2, ids: ['a', 'b'] },
			{ pos: 3, ids: ['a', 'b'] },
			{ pos: 4, ids: ['a', 'b'] },
			{ pos: 5, ids: ['a', 'b'] },
			{ pos: 10, ids: ['a', 'b'] },
			{ pos: 11, ids: [] },
		]);
	});
it('retains outside selections and empty cursors', () => {
	expect(commentSelectionRange(doc, 1, 2)).toEqual({ from: 1, to: 2 });
	expect(commentSelectionRange(doc, 6, 6)).toEqual({ from: 6, to: 6 });
});
