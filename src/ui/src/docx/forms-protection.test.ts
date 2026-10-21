// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument, sectionsOf, type DocumentModel } from 'ooxml-core/docx';
import { splitBlock } from 'prosemirror-commands';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { DocxEditorElement, registerDocxEditor } from './index';
import type { PartsController } from './parts-controller';

registerDocxEditor();

function mount(protection: DocumentModel['protection']) {
	const model = createDocument();
	if (protection) model.protection = protection;
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.documentModel = model;
	document.body.append(editor);
	return (editor as unknown as { core: { parts: PartsController } }).core.parts;
}

const openHeader = (parts: PartsController) => {
	const container = document.createElement('div');
	document.body.append(container);
	return parts.editOnPage({
		container,
		kind: 'headers',
		slot: 'default',
		sectionIndex: 0,
		blocks: [{ type: 'paragraph', id: 'h0', runs: [{ text: 'Header' }] }],
		closed() {},
	});
};

describe('forms protection outside the body', () => {
	afterEach(() => document.body.replaceChildren());

	it('keeps headers and footers closed while forms protection is enforced', () => {
		expect(openHeader(mount({ edit: 'forms', enforced: true }))).toBe(false);
	});

	it('opens them when the document is not protected, or the protection is not enforced', () => {
		expect(openHeader(mount({ edit: 'forms', enforced: false }))).toBe(true);
		document.body.replaceChildren();
		expect(openHeader(mount(undefined))).toBe(true);
	});
});

/** A form protected for filling in, in two sections; `free` names the one left unprotected. */
function sectionedForm(free: 'first' | 'last') {
	const model = createDocument();
	model.blocks = [
		{ type: 'paragraph', id: 'p1', runs: [{ text: 'First section' }] },
		{ type: 'paragraph', id: 'p2', runs: [{ text: 'Last section' }] },
	];
	const [base] = sectionsOf(model);
	model.sections = [
		{ ...structuredClone(base!), endsAtBlockId: 'p1' },
		{ ...structuredClone(base!), endsAtBlockId: 'p2' },
	];
	model.sections[free === 'first' ? 0 : 1]!.formProtection = false;
	model.protection = { edit: 'forms', enforced: true };
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.documentModel = model;
	document.body.append(editor);
	return { editor, view: (editor as unknown as { view: EditorView }).view };
}

/** Presses Enter at `pos`. */
function split(view: EditorView, pos: number): void {
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, pos)));
	splitBlock(view.state, view.dispatch);
}

/** Types `text` at the start of the paragraph holding the selection; the paragraphs' text. */
function typeHere(view: EditorView, text: string): string[] {
	view.dispatch(view.state.tr.insertText(text, view.state.selection.$from.start()));
	const texts: string[] = [];
	view.state.doc.forEach((node) => texts.push(node.textContent));
	return texts;
}

describe('forms protection and sections', () => {
	afterEach(() => document.body.replaceChildren());

	it('lets paragraphs be added at the end of an unprotected last section', () => {
		const { editor, view } = sectionedForm('last');
		split(view, view.state.doc.content.size - 1);
		expect(typeHere(view, 'More')).toEqual(['First section', 'Last section', 'More']);
		expect(editor.documentModel!.blocks.map((block) => block.id)).toHaveLength(3);
		expect(new Set(editor.documentModel!.blocks.map((block) => block.id)).size).toBe(3);
	});

	it('does not move text of an unprotected section into the protected one after it', () => {
		// The break after the first section stays on the first part of a split paragraph.
		const { view } = sectionedForm('first');
		const before = view.state.doc;
		split(view, view.state.doc.child(0).nodeSize - 1);
		expect(view.state.doc.eq(before)).toBe(true);
		split(view, 6);
		expect(view.state.doc.eq(before)).toBe(true);
		expect(typeHere(view, 'Still ')).toEqual(['Still First section', 'Last section']);
	});
});

describe('forms protection and paragraph ids', () => {
	afterEach(() => document.body.replaceChildren());

	it('keeps paragraph ids unique when Enter splits a text field in a protected form', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p1',
				runs: [
					{ text: 'Name: ' },
					{ text: '', fieldChar: 'begin', formField: { kind: 'text', name: 'Name' } },
					{ text: '', fieldCode: ' FORMTEXT ' },
					{ text: '', fieldChar: 'separate' },
					{ text: 'Ada Lovelace', field: { instr: 'FORMTEXT' } },
					{ text: '', fieldChar: 'end' },
				],
			},
		];
		model.protection = { edit: 'forms', enforced: true };
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		editor.documentModel = model;
		document.body.append(editor);
		const view = (editor as unknown as { view: EditorView }).view;
		let inField = -1;
		view.state.doc.descendants((node, pos) => {
			if (node.isText && node.text === 'Ada Lovelace') inField = pos + 3;
		});
		split(view, inField);
		const ids: string[] = [];
		view.state.doc.forEach((node) => ids.push(String(node.attrs.id)));
		expect(ids).toHaveLength(2);
		expect(new Set(ids).size).toBe(2);
	});
});
