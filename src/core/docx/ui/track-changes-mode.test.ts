import { describe, expect, it, vi } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { createCollaborationIdGenerator } from './collaboration-identity';
import { reviewMarks } from './review-schema';
import { collectRevisionRanges } from './review-commands';
import { REMOTE_TRANSACTION_META, trackChangesPlugin } from './track-changes-mode';

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
	it('records the same UTC timestamp for both sides of a replacement', () => {
		const clock = vi.spyOn(Date, 'now').mockReturnValue(Date.UTC(2026, 9, 7, 10, 0, 0));
		try {
			const editor = state();
			const next = editor.apply(editor.tr.insertText('New', 1, 3));
			const marks: { author: string; date: string }[] = [];
			next.doc.descendants((node) => {
				for (const mark of node.marks)
					if (mark.type.name === 'insertion' || mark.type.name === 'deletion')
						marks.push({ author: mark.attrs.author, date: mark.attrs.date });
			});
			expect(marks).toEqual([
				{ author: 'Ada', date: '2026-10-07T10:00:00.000Z' },
				{ author: 'Ada', date: '2026-10-07T10:00:00.000Z' },
			]);
		} finally {
			clock.mockRestore();
		}
	});
	it('keeps more than 100 moves at a fixed time independently named', () => {
		const clock = vi.spyOn(Date, 'now').mockReturnValue(1);
		try {
			const names = Array.from({ length: 101 }, () => {
				const editor = state();
				const tr = editor.tr.delete(1, 2);
				tr.insertText('H', tr.mapping.map(6));
				const ranges = collectRevisionRanges(editor.apply(tr.setMeta('uiEvent', 'drop')).doc);
				expect(ranges).toHaveLength(2);
				expect(ranges[0]!.move).toBe(ranges[1]!.move);
				return ranges[0]!.move;
			});
			expect(new Set(names).size).toBe(101);
		} finally {
			clock.mockRestore();
		}
	});
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
