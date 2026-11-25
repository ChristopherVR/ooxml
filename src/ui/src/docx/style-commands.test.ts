// @vitest-environment jsdom
import { createDocument, loadDocx, saveDocx, type DocumentModel } from 'ooxml-core/docx';
import { redo, undo } from 'prosemirror-history';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { afterEach, expect, it } from 'vitest';
import { DocxEditorElement, registerDocxEditor } from './index';
import { commitStyleEdit, selectionStyle } from './style-commands';
import { must, paragraphAt, tableAt } from './test-support';

registerDocxEditor();
afterEach(() => document.body.replaceChildren());

function mount(model: DocumentModel = createDocument()) {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.documentModel = model;
	document.body.append(editor);
	const view = (editor as unknown as { view: EditorView }).view;
	return { editor, view, current: () => must(editor.documentModel) };
}

function textDocument() {
	const model = createDocument();
	model.blocks = [
		{ type: 'paragraph', id: 'p1', align: 'center', runs: [{ text: 'Some text', bold: true }] },
	];
	return model;
}

it('creates from the selection and applies, modifies, undoes and redoes the style', async () => {
	const { editor, view, current } = mount(textDocument());
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 10)));
	const selection = selectionStyle(view, current());
	expect(selection).toMatchObject({
		runFormatting: { bold: true },
		paragraphFormatting: { align: 'center' },
		paragraphId: 'Normal',
	});
	expect(
		commitStyleEdit(view, current(), {
			kind: 'paragraph',
			name: 'Callout',
			basedOn: 'Normal',
			next: 'Normal',
			paragraphFormatting: selection.paragraphFormatting,
			runFormatting: selection.runFormatting,
		}),
	).toBe(true);
	expect(paragraphAt(current().blocks, 0).style).toBe('Callout');
	expect(current().paragraphStyles?.styles.Callout?.formatting.align).toBe('center');

	expect(
		commitStyleEdit(view, current(), {
			kind: 'paragraph',
			id: 'Callout',
			name: 'Notice',
			basedOn: 'Normal',
			next: 'Normal',
			paragraphFormatting: { align: 'right' },
			runFormatting: { italic: true },
		}),
	).toBe(true);
	expect(current().paragraphStyles?.styles.Callout?.name).toBe('Notice');
	expect(editor.shadowRoot?.querySelector<HTMLElement>('.ProseMirror p')?.style.textAlign).toBe(
		'right',
	);
	expect(undo(view.state, view.dispatch)).toBe(true);
	expect(current().paragraphStyles?.styles.Callout?.name).toBe('Callout');
	expect(undo(view.state, view.dispatch)).toBe(true);
	expect(current().paragraphStyles?.styles.Callout).toBeUndefined();
	expect(paragraphAt(current().blocks, 0)).toMatchObject({
		align: 'center',
		runs: [{ text: 'Some text', bold: true }],
	});
	expect(redo(view.state, view.dispatch)).toBe(true);
	expect(redo(view.state, view.dispatch)).toBe(true);
	const reopened = await loadDocx(await saveDocx(current()));
	expect(reopened.model.paragraphStyles?.styles.Callout?.name).toBe('Notice');
	expect(paragraphAt(reopened.model.blocks, 0).style).toBe('Callout');

	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 10)));
	view.dom.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
	expect(paragraphAt(current().blocks, 1).style).toBe('Normal');
});

it('applies the linked character style to a partial selection and undoes both definitions', () => {
	const { view, current } = mount(textDocument());
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 5)));
	expect(
		commitStyleEdit(view, current(), {
			kind: 'linked',
			name: 'Linked Note',
			paragraphFormatting: {},
			runFormatting: { italic: true },
		}),
	).toBe(true);
	expect(paragraphAt(current().blocks, 0).runs).toMatchObject([
		{ text: 'Some', style: 'LinkedNoteChar' },
		{ text: ' text' },
	]);
	expect(current().paragraphStyles?.styles.LinkedNote?.linkedStyle).toBe('LinkedNoteChar');
	expect(undo(view.state, view.dispatch)).toBe(true);
	expect(current().paragraphStyles?.styles.LinkedNote).toBeUndefined();
	expect(current().characterStyles?.styles.LinkedNoteChar).toBeUndefined();
	expect(paragraphAt(current().blocks, 0).runs).toMatchObject([{ text: 'Some text', bold: true }]);
});

it('creates a table style for the current table and restores its previous reference on undo', async () => {
	const model = createDocument();
	model.blocks = [
		{
			type: 'table',
			id: 't1',
			rows: [
				[
					{
						paragraphs: [{ type: 'paragraph', id: 'p1', runs: [{ text: 'Cell text' }] }],
					},
				],
			],
		},
	];
	const { view, current } = mount(model);
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 4)));
	expect(
		commitStyleEdit(view, current(), {
			kind: 'table',
			name: 'Report Table',
			basedOn: 'TableNormal',
			paragraphFormatting: { align: 'center' },
			runFormatting: { bold: true },
		}),
	).toBe(true);
	expect(tableAt(current().blocks, 0).style).toBe('ReportTable');
	expect(undo(view.state, view.dispatch)).toBe(true);
	expect(tableAt(current().blocks, 0).style).toBeUndefined();
	expect(redo(view.state, view.dispatch)).toBe(true);
	const reopened = await loadDocx(await saveDocx(current()));
	expect(tableAt(reopened.model.blocks, 0).style).toBe('ReportTable');
	expect(reopened.model.tableStyles?.styles.ReportTable?.paragraphFormatting?.align).toBe('center');
});

it('offers style creation from the gallery and modification from the Styles pane', () => {
	const { editor, current } = mount(textDocument());
	const root = must(editor.shadowRoot);
	const button = (parent: ParentNode, title: string) =>
		must(
			[...parent.querySelectorAll<HTMLButtonElement>('button')].find(
				(candidate) =>
					candidate.textContent?.trim() === title || candidate.getAttribute('aria-label') === title,
			),
		);
	const dialog = () => must(root.querySelector('office-ui-dialog.dve-style-dialog'));
	const name = (element: ParentNode) =>
		must(element.querySelector<HTMLInputElement>('[aria-label="Style name"]'));
	button(root, 'More styles').click();
	const gallery = must(
		root.querySelector('.styles-popover') ?? document.querySelector('.styles-popover'),
	);
	button(gallery, 'Create a Style').click();
	const create = dialog();
	name(create).value = 'Callout';
	button(create, 'OK').click();
	expect(current().paragraphStyles?.styles.Callout?.name).toBe('Callout');

	button(root, 'Styles pane').click();
	const modifyButton = button(must(root.querySelector('.dve-styles-pane')), 'Modify Style');
	modifyButton.focus();
	modifyButton.click();
	const modify = dialog();
	expect(name(modify).value).toBe('Callout');
	name(modify).value = 'Notice';
	button(modify, 'OK').click();
	expect(current().paragraphStyles?.styles.Callout?.name).toBe('Notice');
	expect(root.activeElement).toBe(modifyButton);
});
