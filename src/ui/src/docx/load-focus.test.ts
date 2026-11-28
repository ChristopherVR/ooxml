// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { createDocument, saveDocx } from 'ooxml-core/docx';
import { registerDocxEditor } from './index';
import type { DocxEditorElement } from './component';
import type { DocumentView } from './document-views';

registerDocxEditor();
afterEach(() => document.body.replaceChildren());

function documentBytes(): Promise<Uint8Array> {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p1', runs: [{ text: 'A preview' }] }];
	return saveDocx(model);
}

/** A host page with its own focused control next to a surface, as a file list beside a preview. */
function hostPage(readOnly: boolean) {
	const button = document.createElement('button');
	document.body.append(button);
	button.focus();
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.surfaceOnly = true;
	editor.readOnly = readOnly;
	document.body.append(editor);
	return { button, editor };
}

const flush = () => new Promise((resolve) => setTimeout(resolve, 20));

it.each<[DocumentView, boolean]>([
	['print', true],
	['print', false],
	['draft', true],
	['web', true],
	['read', true],
])(
	'keeps the host focus when a %s view is set and a document loads (read-only %s, #380)',
	async (mode, readOnly) => {
		const { button, editor } = hostPage(readOnly);
		expect(editor.getNavigationProps().onViewModeChange(mode)).toBe(true);
		await flush();
		await editor.load(await documentBytes());
		await flush();
		expect(document.activeElement).toBe(button);
		expect(editor.getNavigationProps().viewMode).toBe(mode);
	},
);

it('still hands the focus to the document after a switch made on its own ribbon', async () => {
	const button = document.createElement('button');
	document.body.append(button);
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.readOnly = true;
	document.body.append(editor);
	await editor.load(await documentBytes());
	button.focus();
	const root = editor.shadowRoot!;
	const select = root.querySelector<HTMLSelectElement>('[aria-label="Layout view"]')!;
	// The View tab is open, as when the user picks the view there.
	select.closest<HTMLElement>('[hidden]')!.hidden = false;
	select.focus();
	expect(root.activeElement).toBe(select);
	select.value = 'print';
	select.dispatchEvent(new Event('change'));
	await flush();
	// A read-only document has no editable surface, so its canvas takes the focus.
	expect(root.activeElement).toBe(root.querySelector('.dve-canvas'));
});
