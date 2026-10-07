import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { describe, expect, it } from 'vitest';
import { reviewMarks } from './review-schema.js';
import {
	addComment,
	commentAnchors,
	commentIdsAtSelection,
	deleteComment,
	removeCommentAnchor,
	replyToComment,
	resolveComment,
} from './comment-commands.js';

const schema = new Schema({
	nodes: { doc: { content: 'paragraph+' }, paragraph: { content: 'text*' }, text: {} },
	marks: { comment: reviewMarks.comment },
});
function view(editable = true): EditorView {
	const doc = schema.node('doc', null, [
		schema.node('paragraph', null, schema.text('Alpha')),
		schema.node('paragraph', null, schema.text('Beta')),
	]);
	const host = {
		state: EditorState.create({ doc, selection: TextSelection.create(doc, 2, 10) }),
		editable,
		dispatch(transaction: Transaction) {
			this.state = this.state.apply(transaction);
		},
		focus() {},
	};
	return host as unknown as EditorView;
}

describe('shared Word comment commands', () => {
	it('retains independent anchors when removing one id from a legacy grouped mark', () => {
		const editor = view();
		editor.dispatch(
			editor.state.tr.addMark(2, 10, schema.marks.comment!.create({ ids: ['c1', 'c2'] })),
		);
		addComment(editor, 'Ada', 'Independent', () => 'c3');
		expect(commentIdsAtSelection(editor)).toEqual(['c1', 'c2', 'c3']);
		removeCommentAnchor(editor, 'c1');
		expect(commentIdsAtSelection(editor)).toEqual(['c2', 'c3']);
	});
	it('retains overlapping anchors across paragraphs when one comment is removed', () => {
		const editor = view();
		expect(addComment(editor, 'Ada', 'First', () => 'c1')?.id).toBe('c1');
		expect(addComment(editor, 'Bob', 'Second', () => 'c2')?.id).toBe('c2');
		expect(commentIdsAtSelection(editor)).toEqual(['c1', 'c2']);
		removeCommentAnchor(editor, 'c1');
		expect(commentIdsAtSelection(editor)).toEqual(['c2']);
		expect(commentAnchors(editor)).toEqual([{ id: 'c2', from: 2 }]);
	});
	it('leaves a read-only document untouched', () => {
		const editor = view(false);
		const original = editor.state.doc;
		expect(addComment(editor, 'Ada', 'Blocked')).toBeNull();
		removeCommentAnchor(editor, 'c1');
		expect(editor.state.doc).toBe(original);
	});
	it('keeps reply and resolution operations immutable', () => {
		const original = [{ id: 'c1', author: 'Ada', text: 'Root' }];
		const replied = replyToComment(original, 'c1', 'Bob', 'Reply', () => 'c2');
		expect(original).toHaveLength(1);
		expect(replied[1]).toMatchObject({ id: 'c2', parentId: 'c1', author: 'Bob' });
		const resolved = resolveComment(replied, 'c1', true);
		expect(replied[0]?.resolved).toBeUndefined();
		expect(resolved[0]?.resolved).toBe(true);
		expect(deleteComment(resolved, 'c1')).toEqual([]);
	});
});
