// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createDocument, loadDocx, type DocumentModel, type Paragraph } from 'ooxml-core/docx';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { undo } from 'prosemirror-history';
import { DocxEditorElement, registerDocxEditor } from './index';
import { computeMenuItems, menuStateFromView } from './context-menu-items';

registerDocxEditor();

const view = (editor: DocxEditorElement) => (editor as unknown as { view: EditorView }).view;
const shadow = (editor: DocxEditorElement) => editor.shadowRoot!;
const note = (prefix: string, id: string, kind: 'footnote' | 'endnote', text: string) => ({
	id,
	blocks: [
		{
			type: 'paragraph' as const,
			id: `${prefix}${id}-p0`,
			runs: [{ text: '', noteMark: kind, verticalAlign: 'superscript' as const }, { text }],
		},
	],
});

function notesModel(): DocumentModel {
	const model = createDocument();
	model.blocks = [
		{
			type: 'paragraph',
			id: 'p1',
			runs: [
				{ text: 'Alpha' },
				{ text: '', noteReference: { kind: 'footnote', id: '1' }, verticalAlign: 'superscript' },
				{ text: ' Beta' },
				{ text: '', noteReference: { kind: 'endnote', id: '1' }, verticalAlign: 'superscript' },
			],
		},
	];
	model.footnotes = [note('fn', '1', 'footnote', ' Foot text')];
	model.endnotes = [note('en', '1', 'endnote', ' End text')];
	return model;
}

function mount(model: DocumentModel): DocxEditorElement {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.documentModel = model;
	document.body.append(editor);
	return editor;
}
const references = (editor: DocxEditorElement) =>
	(editor.documentModel!.blocks[0] as Paragraph).runs
		.filter((run) => run.noteReference)
		.map((run) => `${run.noteReference!.kind}:${run.noteReference!.id}`);
const texts = (notes: DocumentModel['footnotes']) =>
	(notes ?? []).map((item) => (item.blocks[0] as Paragraph).runs.map((run) => run.text).join(''));

describe('Convert Notes', () => {
	afterEach(() => document.body.replaceChildren());

	it('swaps footnotes and endnotes from the dialog, and undoes in one step', async () => {
		const editor = mount(notesModel());
		shadow(editor)
			.querySelector<HTMLButtonElement>('.dve-ribbon [aria-label="Convert notes"]')!
			.click();
		const dialog = shadow(editor).querySelector<HTMLElement>('.dve-convert-notes-dialog')!;
		const radio = (label: string) =>
			dialog.querySelector<HTMLInputElement>(`input[aria-label="${label}"]`)!;
		expect(radio('Convert all footnotes to endnotes').checked).toBe(true);
		radio('Swap footnotes and endnotes').checked = true;
		[...dialog.querySelectorAll('button')].find((button) => button.textContent === 'OK')!.click();
		expect(references(editor)).toEqual(['endnote:2', 'footnote:2']);
		expect(texts(editor.documentModel!.footnotes)).toEqual([' End text']);
		expect(texts(editor.documentModel!.endnotes)).toEqual([' Foot text']);
		const marks = [...shadow(editor).querySelectorAll('.ProseMirror sup[data-docx-note-kind]')];
		expect(marks.map((mark) => mark.textContent)).toEqual(['i', '1']);

		const saved = (await loadDocx(await editor.saveBytes())).model;
		expect(texts(saved.footnotes)).toEqual([' End text']);
		expect(texts(saved.endnotes)).toEqual([' Foot text']);
		expect(
			await (
				await JSZip.loadAsync(await editor.saveBytes())
			)
				.file('word/document.xml')!
				.async('string'),
		).toMatch(/<w:endnoteReference w:id="2"\/>.*<w:footnoteReference w:id="2"\/>/);

		const pm = view(editor);
		expect(undo(pm.state, pm.dispatch)).toBe(true);
		expect(references(editor)).toEqual(['footnote:1', 'endnote:1']);
		expect(texts(editor.documentModel!.footnotes)).toEqual([' Foot text']);
	});

	it('disables the choices that have no notes', () => {
		const model = notesModel();
		model.endnotes = [];
		(model.blocks[0] as Paragraph).runs.pop();
		const editor = mount(model);
		shadow(editor)
			.querySelector<HTMLButtonElement>('.dve-ribbon [aria-label="Convert notes"]')!
			.click();
		const dialog = shadow(editor).querySelector<HTMLElement>('.dve-convert-notes-dialog')!;
		const radio = (label: string) =>
			dialog.querySelector<HTMLInputElement>(`input[aria-label="${label}"]`)!;
		expect(radio('Convert all footnotes to endnotes').disabled).toBe(false);
		expect(radio('Convert all endnotes to footnotes').disabled).toBe(true);
		expect(radio('Swap footnotes and endnotes').disabled).toBe(false);
	});

	it('offers Convert to Endnote on a footnote reference and converts only that note', () => {
		const editor = mount(notesModel());
		const pm = view(editor);
		pm.dispatch(pm.state.tr.setSelection(TextSelection.create(pm.state.doc, 6)));
		const items = computeMenuItems(menuStateFromView(pm, false));
		const item = items.find((entry) => entry.id === 'note-convert')!;
		expect(item.label).toBe('Convert to Endnote');
		shadow(editor)
			.querySelector('.dve-ribbon [aria-label="Convert notes"]')!
			.dispatchEvent(
				new CustomEvent('ribbon-action', {
					bubbles: true,
					detail: item.command.kind === 'ribbon' ? item.command.action : undefined,
				}),
			);
		expect(references(editor)).toEqual(['endnote:2', 'endnote:1']);
		expect(editor.documentModel!.footnotes).toEqual([]);
		expect(texts(editor.documentModel!.endnotes)).toEqual([' End text', ' Foot text']);
		const moved = editor.documentModel!.endnotes![1]!.blocks[0] as Paragraph;
		expect(moved.runs[0]).toMatchObject({ noteMark: 'endnote', verticalAlign: 'superscript' });
		pm.dispatch(pm.state.tr.setSelection(TextSelection.create(pm.state.doc, 1)));
		expect(
			computeMenuItems(menuStateFromView(pm, false)).some((entry) => entry.id === 'note-convert'),
		).toBe(false);
	});
});
