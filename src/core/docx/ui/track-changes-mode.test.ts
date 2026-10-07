import { describe, expect, it, vi } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { createCollaborationIdGenerator } from './collaboration-identity.js';
import { reviewMarks } from './review-schema.js';
import { collectRevisionRanges } from './review-commands.js';
import { REMOTE_TRANSACTION_META, trackChangesPlugin } from './track-changes-mode.js';

const schema = new Schema({
	nodes: { doc: { content: 'paragraph+' }, paragraph: { content: 'text*' }, text: {} },
	marks: { insertion: reviewMarks.insertion, deletion: reviewMarks.deletion },
});
function state(author = 'Ada', client?: string) {
	return EditorState.create({
		doc: schema.node('doc', null, schema.node('paragraph', null, schema.text('Hello'))),
		plugins: [
			trackChangesPlugin(
				() => author,
				() => true,
				client ? createCollaborationIdGenerator(client) : undefined,
			),
		],
	});
}

describe('shared revision recording', () => {
	it('records insertions and deletions using the supplied schema', () => {
		let editor = state();
		editor = editor.apply(editor.tr.insertText('!', 6));
		editor = editor.apply(editor.tr.delete(1, 3));
		expect(editor.doc.textContent).toBe('Hello!');
		expect(collectRevisionRanges(editor.doc).map(({ kind, author }) => ({ kind, author }))).toEqual(
			[
				{ kind: 'delete', author: 'Ada' },
				{ kind: 'insert', author: 'Ada' },
			],
		);
	});
	it('removes the author’s pending insertion without recording a deletion', () => {
		let editor = state();
		editor = editor.apply(editor.tr.insertText('!', 6));
		editor = editor.apply(editor.tr.delete(6, 7));
		expect(editor.doc.textContent).toBe('Hello');
		expect(collectRevisionRanges(editor.doc)).toEqual([]);
	});
	it('does not reattribute remote replacements', () => {
		const editor = state();
		const next = editor.apply(editor.tr.insertText('!', 6).setMeta(REMOTE_TRANSACTION_META, true));
		expect(next.doc.textContent).toBe('Hello!');
		expect(collectRevisionRanges(next.doc)).toEqual([]);
	});
	it('keeps two clients and fresh standalone plugins distinct at the same time', () => {
		const clock = vi.spyOn(Date, 'now').mockReturnValue(1);
		try {
			const editors = [state('Ada', 'a'), state('Grace', 'b'), state(), state()];
			const ids = editors.map(
				(editor) => collectRevisionRanges(editor.apply(editor.tr.insertText('!', 6)).doc)[0]!.id,
			);
			expect(new Set(ids).size).toBe(4);
		} finally {
			clock.mockRestore();
		}
	});
});
