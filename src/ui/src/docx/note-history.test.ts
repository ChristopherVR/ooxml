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

it('edits and refreshes a note by kind when both note types use the same id', () => {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	document.body.append(editor);
	try {
		const model = createDocument();
		model.footnotes = [
			{ id: '1', blocks: [{ type: 'paragraph', id: 'fn', runs: [{ text: 'Footnote text' }] }] },
		];
		model.endnotes = [
			{ id: '1', blocks: [{ type: 'paragraph', id: 'en', runs: [{ text: 'Endnote text' }] }] },
		];
		editor.documentModel = model;
		const root = editor.shadowRoot!;
		const footnote = root.querySelector<HTMLElement>(
			'.dve-notes-footnote [data-docx-note-id="1"]',
		)!;
		const endnote = root.querySelector<HTMLElement>('.dve-notes-endnote [data-docx-note-id="1"]')!;
		expect(footnote.textContent).toContain('Footnote text');
		expect(endnote.textContent).toContain('Endnote text');
		endnote.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
		const active = (
			editor as unknown as { core: { parts: { activeView(): EditorView } } }
		).core.parts.activeView();
		expect(active.state.doc.textContent).toBe('Endnote text');
		active.dispatch(
			active.state.tr.insertText('Updated endnote', 1, active.state.doc.content.size - 1),
		);
		expect(editor.documentModel!.endnotes![0]!.blocks[0]).toMatchObject({
			runs: [{ text: 'Updated endnote' }],
		});
		expect(editor.documentModel!.footnotes).toEqual(model.footnotes);
		active.dom.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		const mode = root.querySelector<HTMLSelectElement>('[aria-label="Display for review"]')!;
		mode.value = 'original';
		mode.dispatchEvent(new Event('change', { bubbles: true }));
		expect(footnote.textContent).toContain('Footnote text');
		expect(endnote.textContent).toContain('Updated endnote');
	} finally {
		editor.remove();
	}
});

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
