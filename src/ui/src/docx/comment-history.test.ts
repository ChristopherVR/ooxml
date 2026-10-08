// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { createDocument, loadDocx } from 'ooxml-core/docx';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { DocxEditorElement } from './index';
import './index';
const editors: DocxEditorElement[] = [];
afterEach(() => {
	for (const editor of editors.splice(0)) editor.remove();
});
it('undoes pane threads, replies, resolution and deletion with their exported anchors', async () => {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.documentModel = {
		...createDocument(),
		blocks: [{ type: 'paragraph', id: 'p', runs: [{ text: 'Alpha' }] }],
	};
	document.body.append(editor);
	editors.push(editor);
	const view = (editor as unknown as { view: EditorView }).view;
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 6)));
	const root = editor.shadowRoot!;
	const pane = () => root.querySelector('.dve-comments-panel')!.shadowRoot!;
	const button = (label: string) =>
		[
			...pane().querySelectorAll<HTMLButtonElement>('button'),
			...(root
				.querySelector('office-ui-title-bar')
				?.shadowRoot?.querySelectorAll<HTMLButtonElement>('button') ?? []),
			...root.querySelectorAll<HTMLButtonElement>('button'),
		].find((item) => item.getAttribute('aria-label') === label || item.textContent === label)!;
	root
		.querySelector('office-ui-ribbon-actions')!
		.shadowRoot!.querySelector<HTMLButtonElement>('[part="comments"]')!
		.click();
	pane().querySelector<HTMLTextAreaElement>('[aria-label="New comment"]')!.value = 'Root review';
	button('Add comment').click();
	pane().querySelector<HTMLTextAreaElement>('[aria-label="Reply"]')!.value = 'Reply review';
	button('Reply').click();
	button('Resolve').click();
	button('Delete').click();
	expect(editor.documentModel!.comments ?? []).toEqual([]);
	button('Undo').click();
	expect(editor.documentModel!.comments).toHaveLength(2);
	expect(editor.documentModel!.comments![0]!.resolved).toBe(true);
	expect(pane().textContent).toContain('Reply review');
	const restored = (await loadDocx(await editor.saveBytes())).model;
	expect(restored.comments).toHaveLength(2);
	expect(restored.blocks[0]).toMatchObject({
		runs: [{ text: 'Alpha', commentIds: [restored.comments![0]!.id] }],
	});
	button('Undo').click();
	expect(editor.documentModel!.comments![0]!.resolved).toBe(false);
	button('Undo').click();
	expect(editor.documentModel!.comments).toHaveLength(1);
	expect(pane().textContent).not.toContain('Reply review');
	button('Undo').click();
	expect(editor.documentModel!.comments ?? []).toEqual([]);
	expect(pane().textContent).not.toContain('Root review');
	const undone = (await loadDocx(await editor.saveBytes())).model;
	expect(undone.comments ?? []).toEqual([]);
	expect(undone.blocks[0]).toMatchObject({ runs: [{ text: 'Alpha' }] });
	for (let index = 0; index < 4; index++) button('Redo').click();
	expect(editor.documentModel!.comments ?? []).toEqual([]);
});
