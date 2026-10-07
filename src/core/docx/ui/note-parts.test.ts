import { expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { history, undo, redo } from 'prosemirror-history';
import type { EditorView } from 'prosemirror-view';
import { createDocument, type DocumentModel } from '../model';
import { listRevisions } from '../revision-commands';
import { notePartsJson, restoreNoteParts } from './note-parts';
import { acceptAllChanges, rejectAllChanges, hasAnyChange } from './review-commands';
import { markSpecs } from './schema-marks';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+', attrs: { noteParts: { default: null } } },
		paragraph: { content: 'text*' },
		text: {},
	},
	marks: markSpecs,
});
function source(): DocumentModel {
	return {
		...createDocument(),
		footnotes: [
			{
				id: '1',
				blocks: [
					{
						type: 'paragraph',
						id: 'fn',
						runs: [
							{ text: 'Keep' },
							{ text: 'Added', revision: { id: 'f1', kind: 'insert', author: 'Ada' } },
						],
					},
				],
			},
		],
		endnotes: [
			{
				id: '1',
				blocks: [
					{
						type: 'paragraph',
						id: 'en',
						runs: [
							{ text: 'Removed', revision: { id: 'e1', kind: 'delete', author: 'Bob' } },
							{
								text: 'Keep',
								bold: true,
								revision: {
									id: 'format',
									kind: 'formatChange',
									author: 'Bob',
									previousRunPropertiesXml:
										'<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:b w:val="0"/></w:rPr>',
								},
							},
						],
					},
				],
			},
		],
	};
}
function view(model: DocumentModel, bodyRevision = true): EditorView {
	const editor = {
		editable: true,
		state: EditorState.create({
			doc: schema.node(
				'doc',
				{ noteParts: notePartsJson(model) },
				schema.node(
					'paragraph',
					null,
					schema.text(
						'Body',
						bodyRevision ? [schema.marks.insertion!.create({ id: 'body', author: 'Ada' })] : [],
					),
				),
			),
			plugins: [history()],
		}),
		dispatch(tr: import('prosemirror-state').Transaction) {
			this.state = this.state.apply(tr);
		},
		focus() {},
	};
	return editor as unknown as EditorView;
}
for (const mode of ['accept', 'reject'] as const) {
	it(`resolves both note kinds and the body in one undoable ${mode} command`, () => {
		const model = source();
		const before = structuredClone(model);
		const editor = view(model);
		const initial = editor.state.doc;
		expect((mode === 'accept' ? acceptAllChanges : rejectAllChanges)(editor)).toBe(true);
		const notes = restoreNoteParts(editor.state.doc.attrs.noteParts, {});
		expect(listRevisions({ ...createDocument(), ...notes })).toHaveLength(0);
		const footnote = notes.footnotes![0]!.blocks[0]!;
		const endnote = notes.endnotes![0]!.blocks[0]!;
		if (footnote.type !== 'paragraph' || endnote.type !== 'paragraph')
			throw new Error('Expected paragraphs');
		expect(footnote.runs.map((run) => run.text).join('')).toBe(
			mode === 'accept' ? 'KeepAdded' : 'Keep',
		);
		expect(endnote.runs.map((run) => run.text).join('')).toBe(
			mode === 'accept' ? 'Keep' : 'RemovedKeep',
		);
		expect(endnote.runs.at(-1)?.bold).toBe(mode === 'accept');
		expect(editor.state.doc.textContent).toBe(mode === 'accept' ? 'Body' : '');
		const resolved = editor.state.doc;
		expect(undo(editor.state, editor.dispatch.bind(editor))).toBe(true);
		expect(editor.state.doc.eq(initial)).toBe(true);
		expect(redo(editor.state, editor.dispatch.bind(editor))).toBe(true);
		expect(editor.state.doc.eq(resolved)).toBe(true);
		expect(model).toEqual(before);
	});
}
it('enables Accept All for note-only changes and blocks read-only commands', () => {
	const editor = view(source(), false);
	expect(hasAnyChange(editor)).toBe(true);
	const initial = editor.state.doc;
	Object.defineProperty(editor, 'editable', { value: false });
	expect(acceptAllChanges(editor)).toBe(false);
	expect(editor.state.doc).toBe(initial);
});
it('leaves the whole document unchanged when a note paragraph cannot be merged', () => {
	const model = source();
	const note = model.footnotes![0]!.blocks[0]!;
	if (note.type !== 'paragraph') throw new Error('Expected paragraph');
	note.markRevision = { id: 'mark', kind: 'insert', author: 'Ada' };
	const editor = view(model);
	const initial = editor.state.doc;
	expect(() => rejectAllChanges(editor)).toThrow('no following paragraph');
	expect(editor.state.doc).toBe(initial);
	expect(undo(editor.state, editor.dispatch.bind(editor))).toBe(false);
});
