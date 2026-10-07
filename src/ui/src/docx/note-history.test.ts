// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { createDocument, loadDocx } from 'ooxml-core/docx';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history, undo, redo } from 'prosemirror-history';
import { notePartsJson, NotePartsStep } from 'ooxml-core/docx/ui';
import { modelToDoc, docToModel } from './model-adapter';
import { insertNote } from './note-commands';
import { DocxEditorElement } from './index';
import './index';

for (const kind of ['footnote', 'endnote'] as const)
	it(`inserts the ${kind} reference and content into one document history event`, () => {
		let model = createDocument();
		const initial = docToModel(modelToDoc(model), model);
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc: modelToDoc(model), plugins: [history()] }),
			dispatchTransaction(tr) {
				view.updateState(view.state.apply(tr));
				model = docToModel(view.state.doc, model);
			},
		});
		try {
			const key = kind === 'footnote' ? 'footnotes' : 'endnotes';
			insertNote(view, model, kind);
			expect(model[key]).toHaveLength(1);
			expect(model.blocks[0]).toMatchObject({ runs: [{ noteReference: { kind, id: '1' } }] });
			const inserted = structuredClone(model);
			expect(undo(view.state, view.dispatch)).toBe(true);
			expect(model).toEqual(initial);
			expect(redo(view.state, view.dispatch)).toBe(true);
			expect(model).toEqual(inserted);
		} finally {
			view.destroy();
		}
	});

it('refreshes note previews after body history and saves the current snapshot', async () => {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	document.body.append(editor);
	try {
		const model = createDocument();
		model.footnotes = [
			{ id: '1', blocks: [{ type: 'paragraph', id: 'fn', runs: [{ text: 'First note' }] }] },
		];
		editor.documentModel = model;
		const view = (editor as unknown as { view: EditorView }).view;
		const next = structuredClone(model);
		next.footnotes![0]!.blocks = [
			{ type: 'paragraph', id: 'fn', runs: [{ text: 'Updated note' }] },
		];
		view.dispatch(view.state.tr.step(new NotePartsStep(notePartsJson(next))));
		expect(editor.shadowRoot!.querySelector('.dve-note-body')!.textContent).toBe('Updated note');
		expect((await loadDocx(await editor.saveBytes())).model.footnotes![0]!.blocks[0]).toMatchObject(
			{ runs: [{ text: 'Updated note' }] },
		);
		expect(undo(view.state, view.dispatch)).toBe(true);
		expect(editor.shadowRoot!.querySelector('.dve-note-body')!.textContent).toBe('First note');
	} finally {
		editor.remove();
	}
});
