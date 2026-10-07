import { expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection, NodeSelection, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { history, undo, redo } from 'prosemirror-history';
import {
	commentHistoryAttributes,
	commentThreadsFromDoc,
	updateCommentThreads,
} from './comment-history';
import {
	addComment,
	removeCommentThread,
	replyToComment,
	resolveComment,
	commentIdsAtSelection,
} from './comment-commands';
import { imageNodeSpec } from './inline-content-schema';
import { markSpecs } from './schema-marks';
const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+', attrs: commentHistoryAttributes },
		paragraph: { content: 'inline*' },
		text: { group: 'inline' },
		image: imageNodeSpec,
	},
	marks: markSpecs,
});
for (const picture of [false, true])
	it(`undoes ${picture ? 'picture' : 'text'} comment creation, reply, resolution and deletion together with anchors`, () => {
		const atom = picture ? schema.node('image', { relId: 'rId1' }) : schema.text('Alpha');
		const doc = schema.node(
			'doc',
			{ commentThreads: '[]' },
			schema.node('paragraph', null, [atom]),
		);
		const host = {
			state: EditorState.create({
				doc,
				plugins: [history()],
				selection: picture ? NodeSelection.create(doc, 1) : TextSelection.create(doc, 1, 6),
			}),
			editable: true,
			dispatch(tr: Transaction) {
				this.state = this.state.applyTransaction(tr).state;
			},
		};
		const view = host as unknown as EditorView;
		const threads = () => commentThreadsFromDoc(host.state.doc)!;
		const dispatch = (tr: Transaction) => host.dispatch(tr);
		expect(addComment(view, 'Ada', 'Root', () => 'a')).not.toBeNull();
		expect(
			updateCommentThreads(
				view,
				replyToComment(threads(), 'a', 'Bob', 'Reply', () => 'b'),
			),
		).toBe(true);
		expect(updateCommentThreads(view, resolveComment(threads(), 'a', true))).toBe(true);
		expect(removeCommentThread(view, 'a')).toBe(true);
		expect(threads()).toEqual([]);
		expect(commentIdsAtSelection(view)).toEqual([]);
		expect(undo(host.state, dispatch)).toBe(true);
		expect(threads()).toHaveLength(2);
		expect(threads()[0]!.resolved).toBe(true);
		expect(commentIdsAtSelection(view)).toEqual(['a']);
		expect(undo(host.state, dispatch)).toBe(true);
		expect(threads()[0]!.resolved).toBe(false);
		expect(undo(host.state, dispatch)).toBe(true);
		expect(threads()).toHaveLength(1);
		expect(undo(host.state, dispatch)).toBe(true);
		expect(threads()).toEqual([]);
		expect(commentIdsAtSelection(view)).toEqual([]);
		for (let index = 0; index < 4; index++) expect(redo(host.state, dispatch)).toBe(true);
		expect(threads()).toEqual([]);
		expect(host.state.doc.textContent).toBe(doc.textContent);
	});
