import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { describe, expect, it } from 'vitest';
import { reviewMarks } from './review-schema';
import { imageNodeSpec } from './inline-content-schema';
import {
	hardBreakNodeSpec,
	pageBreakNodeSpec,
	noteReferenceNodeSpec,
	fieldMarkerNodeSpec,
} from './break-note-schema';
import { runToInlineNodes } from './run-adapter';
import type { TextRun } from '../model';
import {
	addComment,
	commentAnchors,
	commentIdsAtSelection,
	deleteComment,
	removeCommentAnchor,
	replyToComment,
	resolveComment,
} from './comment-commands';

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
	for (const atom of [
		{ text: '\n' },
		{ text: '', break: 'page' },
		{ text: '', noteReference: { kind: 'footnote', id: '1' } },
		{ text: '', fieldChar: 'begin' },
		{
			text: '',
			image: {
				relId: 'rId1',
				partName: 'word/media/a.png',
				contentType: 'image/png',
				widthPx: 10,
				heightPx: 20,
			},
		},
	] satisfies TextRun[])
		it(`finds imported ${atom.text ? 'hard-break' : Object.keys(atom)[1]} anchors together with overlapping marks`, () => {
			const inlineSchema = new Schema({
				nodes: {
					doc: { content: 'paragraph+' },
					paragraph: { content: 'inline*' },
					text: { group: 'inline' },
					image: imageNodeSpec,
					hardBreak: hardBreakNodeSpec,
					pageBreak: pageBreakNodeSpec,
					noteReference: noteReferenceNodeSpec,
					fieldMarker: fieldMarkerNodeSpec,
				},
				marks: { comment: reviewMarks.comment },
			});
			const node = runToInlineNodes({ ...atom, commentIds: ['c2', 'c1', 'c1'] }, inlineSchema)[0]!;
			const doc = inlineSchema.node(
				'doc',
				null,
				inlineSchema.node('paragraph', null, [
					inlineSchema.text('L'),
					node.mark([inlineSchema.marks.comment!.create({ ids: ['c2', 'c3'] })]),
					inlineSchema.text('R'),
				]),
			);
			const host = {
				state: EditorState.create({ doc, selection: TextSelection.create(doc, 2, 3) }),
			};
			const editor = host as unknown as EditorView;
			expect(commentIdsAtSelection(editor)).toEqual(['c1', 'c2', 'c3']);
			expect(commentAnchors(editor)).toEqual([
				{ id: 'c1', from: 2 },
				{ id: 'c2', from: 2 },
				{ id: 'c3', from: 2 },
			]);
		});
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
