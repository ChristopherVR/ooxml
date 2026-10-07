import { describe, expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { history, undo, redo } from 'prosemirror-history';
import { trackChangesPlugin } from './track-changes-mode.js';
import { reviewMarks } from './review-schema.js';
import {
	acceptAllChanges,
	acceptRevisionRange,
	collectRevisionRanges,
	rejectRevisionRange,
} from './review-commands.js';

const schema = new Schema({
	nodes: { doc: { content: 'paragraph+' }, paragraph: { content: 'text*' }, text: {} },
	marks: { insertion: reviewMarks.insertion, deletion: reviewMarks.deletion },
});
function view(editable = true, recording = false): EditorView {
	const doc = schema.node(
		'doc',
		null,
		schema.node('paragraph', null, [
			schema.text('Ada', [schema.marks.insertion!.create({ id: 'legacy', author: 'Ada' })]),
			schema.text('Grace', [schema.marks.insertion!.create({ id: 'legacy', author: 'Grace' })]),
		]),
	);
	const host = {
		state: EditorState.create({
			doc,
			plugins: recording
				? [
						history(),
						trackChangesPlugin(
							() => 'Ada',
							() => true,
						),
					]
				: [],
		}),
		editable,
		dispatch(tr: Transaction) {
			this.state = this.state.apply(tr);
		},
		focus() {},
	};
	return host as unknown as EditorView;
}
describe('shared revision resolution', () => {
	it('undoes and redoes resolution without changing the original author', () => {
		const editor = view(true, true);
		const before = editor.state.doc;
		rejectRevisionRange(editor, collectRevisionRanges(before)[1]!);
		expect(editor.state.doc.textContent).toBe('Ada');
		expect(undo(editor.state, (tr) => editor.dispatch(tr))).toBe(true);
		expect(editor.state.doc.eq(before)).toBe(true);
		expect(redo(editor.state, (tr) => editor.dispatch(tr))).toBe(true);
		expect(editor.state.doc.textContent).toBe('Ada');
	});
	it('keeps legacy colliding IDs from different authors separate', () => {
		const editor = view();
		expect(collectRevisionRanges(editor.state.doc).map((range) => range.author)).toEqual([
			'Ada',
			'Grace',
		]);
		rejectRevisionRange(editor, collectRevisionRanges(editor.state.doc)[0]!);
		expect(editor.state.doc.textContent).toBe('Grace');
		expect(collectRevisionRanges(editor.state.doc)[0]!.author).toBe('Grace');
	});
	it('accepts one author without accepting a colliding peer revision', () => {
		const editor = view();
		acceptRevisionRange(editor, collectRevisionRanges(editor.state.doc)[0]!);
		expect(editor.state.doc.textContent).toBe('AdaGrace');
		expect(collectRevisionRanges(editor.state.doc).map((range) => range.author)).toEqual(['Grace']);
	});
	it('blocks direct revision mutations in a read-only view', () => {
		const editor = view(false);
		const before = editor.state.doc;
		const range = collectRevisionRanges(before)[0]!;
		acceptRevisionRange(editor, range);
		rejectRevisionRange(editor, range);
		expect(acceptAllChanges(editor)).toBe(false);
		expect(editor.state.doc).toBe(before);
	});
	it('ignores a stale revision range after it has already been resolved', () => {
		const editor = view();
		const range = collectRevisionRanges(editor.state.doc)[0]!;
		rejectRevisionRange(editor, range);
		rejectRevisionRange(editor, range);
		expect(editor.state.doc.textContent).toBe('Grace');
		expect(collectRevisionRanges(editor.state.doc)[0]!.author).toBe('Grace');
	});
});
