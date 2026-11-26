// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument, loadDocx, type DocumentModel, type Paragraph } from 'ooxml-core/docx';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { DocxEditorElement, registerDocxEditor } from './index';
import { must } from './test-support';

registerDocxEditor();

function mount(model: DocumentModel): DocxEditorElement {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.documentModel = model;
	document.body.append(editor);
	return editor;
}
const view = (editor: DocxEditorElement) => (editor as unknown as { view: EditorView }).view;
const firstParagraph = (editor: DocxEditorElement) => editor.documentModel!.blocks[0] as Paragraph;
const click = (editor: DocxEditorElement, label: string) =>
	editor.shadowRoot?.querySelector<HTMLButtonElement>(`[aria-label="${label}"]`)?.click();

function withText(text: string, trackChanges = false): DocumentModel {
	const model = createDocument();
	model.blocks[0] = { type: 'paragraph', id: 'p1', runs: [{ text }] };
	model.trackChanges = trackChanges;
	return model;
}

describe('review editing end to end', () => {
	afterEach(() => document.body.replaceChildren());

	it('records typing as a tracked insertion and accepts it', () => {
		const editor = mount(withText('Hello', true));
		const pm = view(editor);
		pm.dispatch(pm.state.tr.insertText(' world', 6));
		expect(pm.state.doc.textContent).toBe('Hello world');
		const runs = firstParagraph(editor).runs;
		expect(runs.map((run) => run.text).join('')).toBe('Hello world');
		expect(runs.find((run) => run.text === ' world')?.revision?.kind).toBe('insert');
		click(editor, 'Accept all');
		const accepted = firstParagraph(editor).runs;
		expect(accepted.map((run) => run.text).join('')).toBe('Hello world');
		expect(accepted.some((run) => run.revision)).toBe(false);
	});

	it('keeps deleted text visible as a tracked deletion and rejects it', () => {
		const editor = mount(withText('Hello world', true));
		const pm = view(editor);
		pm.dispatch(pm.state.tr.delete(6, 12));
		expect(pm.state.doc.textContent).toBe('Hello world');
		const deleted = firstParagraph(editor).runs.find((run) => run.revision?.kind === 'delete');
		expect(deleted?.text).toBe(' world');
		click(editor, 'Reject all');
		const rejected = firstParagraph(editor).runs;
		expect(rejected.map((run) => run.text).join('')).toBe('Hello world');
		expect(rejected.some((run) => run.revision)).toBe(false);
	});

	it('removes the current author’s own pending insertion outright', () => {
		const editor = mount(withText('Hi', true));
		const pm = view(editor);
		pm.dispatch(pm.state.tr.insertText('!', 3));
		pm.dispatch(pm.state.tr.delete(3, 4));
		expect(pm.state.doc.textContent).toBe('Hi');
		expect(firstParagraph(editor).runs.some((run) => run.revision)).toBe(false);
	});

	it('toggles track changes from the Review ribbon', () => {
		const editor = mount(withText('Text'));
		click(editor, 'Track changes');
		expect(editor.documentModel!.trackChanges).toBe(true);
		click(editor, 'Track changes');
		expect(editor.documentModel!.trackChanges).toBe(false);
	});

	it.each([
		{ tab: 'Insert', from: 1, to: 7, text: 'Review' },
		{ tab: 'Insert', from: 4, to: 4, text: 'Review' },
		{ tab: 'Review', from: 7, to: 7, text: '' },
	])(
		'adds an anchored comment from $tab at $from-$to and saves it',
		async ({ tab, from, to, text }) => {
			const editor = mount(withText('Review me'));
			const pm = view(editor);
			pm.dispatch(pm.state.tr.setSelection(TextSelection.create(pm.state.doc, from, to)));
			must(
				editor.shadowRoot?.querySelector<HTMLButtonElement>(
					`.ribbon-panel[data-panel="${tab}"] button[aria-label="Add comment"]`,
				),
			).click();
			const pane = editor.shadowRoot!.querySelector('.dve-comments-panel')!.shadowRoot!;
			const input = pane.querySelector<HTMLTextAreaElement>('[aria-label="New comment"]')!;
			expect(input).toBeTruthy();
			input.value = 'Please check';
			input.dispatchEvent(new Event('input'));
			const submit = [...pane.querySelectorAll<HTMLButtonElement>('button')].find(
				(button) => button.textContent === 'Add comment',
			);
			submit?.click();
			const comments = editor.documentModel!.comments ?? [];
			expect(comments).toHaveLength(1);
			const [comment] = comments;
			expect(must(comment).text).toBe('Please check');
			const anchored = firstParagraph(editor).runs.find((run) =>
				run.commentIds?.includes(must(comment).id),
			);
			expect(anchored?.text).toBe(text);
			expect(pm.state.doc.textContent).toBe('Review me');
			expect(
				firstParagraph(editor)
					.runs.map((run) => run.text)
					.join(''),
			).toBe('Review me');

			const blob = await editor.save();
			const bytes = await new Promise<Uint8Array>((resolve, reject) => {
				const reader = new FileReader();
				reader.onload = () => {
					if (reader.result instanceof ArrayBuffer) resolve(new Uint8Array(reader.result));
					else reject(new Error('Expected DOCX bytes'));
				};
				reader.onerror = () => reject(reader.error);
				reader.readAsArrayBuffer(blob);
			});
			const { model } = await loadDocx(bytes);
			const savedComment = must(model.comments?.[0]);
			const paragraph = model.blocks[0];
			expect(savedComment.text).toBe('Please check');
			expect(
				paragraph?.type === 'paragraph' &&
					paragraph.runs.find((run) => run.commentIds?.includes(savedComment.id))?.text,
			).toBe(text);
			if (text === '') {
				expect(paragraph?.type === 'paragraph' && paragraph.runs.map((run) => run.text)).toEqual([
					'Review',
					'',
					' me',
				]);
				click(editor, 'Delete comment');
				expect(editor.documentModel!.comments ?? []).toEqual([]);
				expect(firstParagraph(editor).runs).toEqual([{ text: 'Review me' }]);
			}
		},
	);
});
