// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { createDocument, saveDocx, type Paragraph } from 'ooxml-core/docx';
import { TextSelection } from 'prosemirror-state';
import { type EditorView } from 'prosemirror-view';
import { DocxEditorElement } from './index';
import './index';
import { withBlankHeaderFooter } from './header-footer-commands';
import { renderBlocks } from './header-footer-view';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const paragraph = (id: string): Paragraph => ({
	type: 'paragraph',
	id,
	align: 'right',
	formatRevision: {
		id: 'pf',
		kind: 'paragraphChange',
		author: 'Ada',
		previousParagraphPropertiesXml: `<w:pPr xmlns:w="${ns}"><w:jc w:val="center"/><w:spacing w:before="240"/></w:pPr>`,
	},
	runs: [
		{
			text: 'Format',
			bold: true,
			formatRevision: {
				id: 'rf',
				kind: 'formatChange',
				author: 'Ada',
				previousRunPropertiesXml: `<w:rPr xmlns:w="${ns}"><w:i/></w:rPr>`,
			},
		},
		{ text: 'Inserted', revision: { id: 'i', kind: 'insert', author: 'Ada' } },
		{ text: 'Deleted', revision: { id: 'd', kind: 'delete', author: 'Ada' } },
	],
});
let editor: DocxEditorElement | undefined;
afterEach(() => {
	editor?.remove();
	editor = undefined;
});

it('projects story previews, including table cells, without mutating source properties', () => {
	const model = createDocument();
	const source = paragraph('cell');
	const blocks = [{ type: 'table' as const, id: 'table', rows: [[{ paragraphs: [source] }]] }];
	const original = structuredClone(blocks);
	const host = document.createElement('div');
	host.append(renderBlocks(blocks, model, 'original'));
	expect(host.textContent).toBe('FormatDeleted');
	expect(host.querySelector('p')!.style.textAlign).toBe('center');
	expect(host.querySelector('p')!.style.marginTop).toBe('16px');
	expect(host.querySelector('strong')).toBeNull();
	expect(host.querySelector('em')).not.toBeNull();
	expect(blocks).toEqual(original);
	host.replaceChildren(renderBlocks(blocks, model, 'final'));
	expect(host.textContent).toBe('FormatInserted');
	expect(host.querySelector('p')!.style.textAlign).toBe('right');
	expect(host.querySelector('strong')).not.toBeNull();
});

it('refreshes active headers, footers and notes while preserving selection and pending revisions', async () => {
	let model = withBlankHeaderFooter(createDocument(), 'headers', () => 'h');
	model = withBlankHeaderFooter(model, 'footers', () => 'f');
	model.sections![0]!.headers!.default!.blocks = [paragraph('header')];
	model.sections![0]!.footers!.default!.blocks = [paragraph('footer')];
	model.footnotes = [{ id: '1', blocks: [paragraph('note')] }];
	editor = document.createElement('docx-editor') as DocxEditorElement;
	document.body.append(editor);
	await editor.load(await saveDocx(model));
	const root = editor.shadowRoot!;
	const source = structuredClone(editor.documentModel);
	const mode = root.querySelector<HTMLSelectElement>('[aria-label="Display for review"]')!;
	const setMode = (value: string) => {
		mode.value = value;
		mode.dispatchEvent(new Event('change', { bubbles: true }));
	};
	for (const selector of [
		'.dve-header [data-slot="default"]',
		'.dve-footer [data-slot="default"]',
		'[data-docx-note-id="1"]',
	]) {
		const slot = root.querySelector<HTMLElement>(selector)!;
		slot.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
		const view = (
			editor as unknown as { core: { parts: { activeView(): EditorView } } }
		).core.parts.activeView();
		const doc = view.state.doc;
		view.dispatch(view.state.tr.setSelection(TextSelection.create(doc, 3)));
		const selection = view.state.selection;
		setMode('original');
		expect(view.state.doc).toBe(doc);
		expect(view.state.selection.eq(selection)).toBe(true);
		expect(view.dom.querySelector('p')!.style.textAlign).toBe('center');
		expect(view.dom.querySelector('strong')!.style.fontWeight).toBe('inherit');
		expect(editor.documentModel).toEqual(source);
		view.dom.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		expect(slot.querySelector('.ProseMirror')).toBeNull();
		expect(slot.querySelector('p')!.style.textAlign).toBe('center');
		expect(slot.textContent).toContain('FormatDeleted');
		expect(slot.textContent).not.toContain('Inserted');
		setMode('final');
		expect(slot.querySelector('p')!.style.textAlign).toBe('right');
		expect(slot.textContent).toContain('FormatInserted');
		expect(slot.textContent).not.toContain('Deleted');
	}
	expect(editor.documentModel).toEqual(source);
});
