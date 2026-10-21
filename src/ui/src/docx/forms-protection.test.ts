// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument, type DocumentModel } from 'ooxml-core/docx';
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
