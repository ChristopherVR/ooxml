// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { createDocument, resolveNumberingLevel, type DocumentModel } from 'ooxml-core/docx';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { docToModel, modelToDoc } from './model-adapter';
import { wordPastePlugin } from './word-paste';

/** HTML Format data Word 16 put on the clipboard, as captured for the core reader's tests. */
const fixture = (name: string): string =>
	readFileSync(
		path.join(import.meta.dirname, `../../../core/docx/__fixtures__/clipboard/${name}.html`),
		'utf8',
	);

const views: EditorView[] = [];
afterEach(() => {
	for (const view of views.splice(0)) view.destroy();
	document.body.innerHTML = '';
	vi.unstubAllGlobals();
});

function editor(model: DocumentModel = createDocument()) {
	vi.stubGlobal('ClipboardEvent', Event);
	const host = document.createElement('div');
	document.body.append(host);
	const doc = modelToDoc(model);
	const view = new EditorView(host, {
		state: EditorState.create({ doc, plugins: [wordPastePlugin(() => model)] }),
	});
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1)));
	views.push(view);
	return { view, model };
}

const paragraphs = (view: EditorView) => {
	const found: Array<{ text: string; attrs: Record<string, unknown> }> = [];
	view.state.doc.descendants((node) => {
		if (node.type.name === 'paragraph') found.push({ text: node.textContent, attrs: node.attrs });
	});
	return found;
};

describe('pasting from Word', () => {
	it('pastes mso-list paragraphs as one real list with its levels', () => {
		const { view, model } = editor();
		view.pasteHTML(fixture('word-numbered'));
		// Pasted into the empty paragraph, the list replaces it rather than following it.
		expect(paragraphs(view)[0]!.text).toBe('Step one');
		const items = paragraphs(view).filter((p) => p.text);
		expect(items.map((p) => [p.text, p.attrs.ilvl])).toEqual([
			['Step one', 0],
			['Sub step a', 1],
			['Sub sub step', 2],
			['Sub step b', 1],
			['Step two', 0],
		]);
		const numId = items[0]!.attrs.numId as number;
		expect(items.every((p) => p.attrs.numId === numId)).toBe(true);
		const catalog = model.numberingCatalog!;
		expect(resolveNumberingLevel(catalog, String(numId), 1)).toMatchObject({
			numFmt: 'lowerLetter',
			lvlText: '%2)',
		});
		// No marker text reaches the document: the list draws its own numbers.
		expect(items[0]!.text).toBe('Step one');
	});

	it('gives each paste its own list', () => {
		const { view } = editor();
		view.pasteHTML(fixture('word-bullets'));
		const first = paragraphs(view).find((p) => p.text === 'First bullet')!.attrs.numId;
		view.pasteHTML(fixture('word-bullets'));
		const ids = new Set(
			paragraphs(view)
				.filter((p) => p.text === 'First bullet')
				.map((p) => p.attrs.numId),
		);
		expect(ids.size).toBe(2);
		expect(ids.has(first)).toBe(true);
	});

	it('applies the styles the document defines and keeps Word-only run formatting', () => {
		const { view, model } = editor();
		view.pasteHTML(fixture('word-styles'));
		const byText = new Map(paragraphs(view).map((p) => [p.text, p.attrs.style]));
		expect(byText.get('Clipboard Heading')).toBe('Heading1');
		expect(byText.get('Second level heading')).toBe('Heading2');
		// The new document has no Quote or Callout Text style: they paste as Normal text.
		expect(byText.get('A quoted paragraph.')).toBe('');
		const runs = docToModel(view.state.doc, model).blocks.flatMap((block) =>
			block.type === 'paragraph' ? block.runs : [],
		);
		const run = (text: string) => runs.find((r) => r.text === text);
		expect(run('bold')).toMatchObject({ bold: true });
		expect(run('small caps')).toMatchObject({ smallCaps: true });
		expect(run('double struck')).toMatchObject({ doubleStrike: true, highlight: 'yellow' });
		expect(run('double struck')?.strike).toBeUndefined();
		expect(run('double underline')).toMatchObject({ underline: true, underlineStyle: 'double' });
		expect(run('red')).toMatchObject({ color: '#ff0000' });
	});

	it('keeps a field as a field', () => {
		const { view, model } = editor();
		view.pasteHTML(fixture('word-table-notes'));
		const runs = docToModel(view.state.doc, model).blocks.flatMap((block) =>
			block.type === 'paragraph' ? block.runs : [],
		);
		expect(runs.find((r) => r.text === '1' && r.field)?.field).toMatchObject({
			instr: ' PAGE ',
		});
	});

	it('leaves HTML from elsewhere to the normal parser', () => {
		const { view, model } = editor();
		view.pasteHTML('<p>Web <b>text</b></p><ul><li>not Word</li></ul>');
		expect(model.numberingCatalog).toBeUndefined();
		expect(paragraphs(view).map((p) => p.text)).toContain('Web text');
	});
});
