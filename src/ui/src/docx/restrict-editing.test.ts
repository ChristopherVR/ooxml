// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument, enforceProtection, type DocumentModel } from 'ooxml-core/docx';
import type { EditorView } from 'prosemirror-view';
import { DocxEditorElement, registerDocxEditor } from './index';

registerDocxEditor();

type Core = { view: EditorView; formatDialogs: { open(kind: 'restrictEditing'): void } };

function mount(protection: DocumentModel['protection']) {
	const model = createDocument();
	model.blocks = [
		{
			type: 'paragraph',
			id: 'p1',
			runs: [{ text: 'hello world' }],
			permissions: [{ id: '1', group: 'everyone', from: 6, to: 11 }],
		},
		{ type: 'paragraph', id: 'p2', runs: [{ text: 'locked' }] },
	];
	if (protection) model.protection = protection;
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.documentModel = model;
	document.body.append(editor);
	const core = (editor as unknown as { core: Core }).core;
	return { editor, core, view: core.view };
}

const texts = (view: EditorView) => {
	const result: string[] = [];
	view.state.doc.forEach((node) => result.push(node.textContent));
	return result;
};

describe('Restrict Editing', () => {
	afterEach(() => document.body.replaceChildren());

	it('lets a read only document change only in its editable regions, and highlights them', () => {
		const { view, editor } = mount({ edit: 'readOnly', enforced: true });
		view.dispatch(view.state.tr.insertText('X', 2));
		expect(texts(view)).toEqual(['hello world', 'locked']);
		view.dispatch(view.state.tr.insertText('big ', 7));
		expect(texts(view)).toEqual(['hello big world', 'locked']);
		expect(editor.shadowRoot!.querySelector('.dve-editable-region')?.textContent).toBe('big world');
	});

	it('does not restrict or highlight a document without protection', () => {
		const { view, editor } = mount(undefined);
		view.dispatch(view.state.tr.insertText('X', 2));
		expect(texts(view)[0]).toBe('hXello world');
		expect(editor.shadowRoot!.querySelector('.dve-editable-region')).toBeNull();
	});

	it('forces Track Changes on under tracked changes protection', () => {
		const { view } = mount({ edit: 'trackedChanges', enforced: true });
		view.dispatch(view.state.tr.insertText('X', 14));
		let marked = false;
		view.state.doc.descendants((node) => {
			if (node.marks.some((mark) => mark.type.name === 'insertion')) marked = true;
		});
		expect(marked).toBe(true);
	});

	it('starts and stops protection with a password from the Review tab', () => {
		const { core, editor } = mount(undefined);
		const review = editor.shadowRoot!.querySelector('[data-panel="Review"]')!;
		expect(review.querySelector('[aria-label="Restrict editing"]')).not.toBeNull();
		core.formatDialogs.open('restrictEditing');
		const root = editor.shadowRoot!;
		const check = (label: string) =>
			root.querySelector<HTMLInputElement>(`input[aria-label="${label}"]`)!;
		const restrict = check('Allow only this type of editing in the document');
		restrict.checked = true;
		restrict.dispatchEvent(new Event('change'));
		const [first, again] = Array.from(
			root.querySelectorAll<HTMLInputElement>('input[type="password"]'),
		);
		first!.value = again!.value = 'pw';
		const click = (text: string) =>
			Array.from(root.querySelectorAll('button'))
				.find((button) => button.textContent === text)!
				.click();
		click('Yes, start enforcing protection');
		const model = () => (editor as unknown as { documentModel: DocumentModel }).documentModel;
		expect(model().protection).toMatchObject({ edit: 'readOnly', enforced: true });
		expect(model().protection?.password).toBeDefined();

		core.formatDialogs.open('restrictEditing');
		const secret = root.querySelector<HTMLInputElement>('input[type="password"]')!;
		secret.value = 'wrong';
		click('Stop protection');
		expect(model().protection?.enforced).toBe(true);
		secret.value = 'pw';
		click('Stop protection');
		expect(model().protection).toEqual({ edit: 'readOnly', enforced: false });
		expect(enforceProtection('readOnly').enforced).toBe(true);
	});
});
