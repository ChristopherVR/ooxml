import { afterEach, expect, it } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { wordAddInSelection } from './add-in-selection';
import { loadWordAddInSelection } from './add-in-selection-loader';
import { schema } from './schema';

let view: EditorView | undefined;
afterEach(() => view?.destroy());

it('reads a read-only selection and writes multiline plain text as paragraphs when editable', () => {
	const paragraph = schema.node('paragraph', { style: 'Custom', keepNext: true }, [
		schema.text('original'),
	]);
	const doc = schema.node('doc', null, [paragraph]);
	view = new EditorView(document.body, {
		state: EditorState.create({ doc, selection: TextSelection.create(doc, 1, 9) }),
		editable: () => false,
	});
	const selection = wordAddInSelection(() => view);
	expect(selection.readText()).toBe('original');
	expect(selection.writeText('blocked')).toBe(false);
	view.setProps({ editable: () => true });
	expect(selection.writeText('first\n\nlast')).toBe(true);
	expect(view.state.doc.textBetween(0, view.state.doc.content.size, '\n')).toBe('first\n\nlast');
	expect(view.state.doc.childCount).toBe(3);
	view.state.doc.forEach((child) => expect(child.attrs).toEqual(paragraph.attrs));
});

it('gives inserted paragraphs the formatting but not the identity of the paragraph they land in', () => {
	const paragraph = schema.node(
		'paragraph',
		{
			id: 'p-source',
			style: 'ListParagraph',
			numId: 3,
			ilvl: 0,
			listLabelText: '1.',
			markRevision: { kind: 'insert', author: 'A' },
			formatRevision: { author: 'B' },
			sourceParagraphPropertiesXml: '<w:pPr/>',
			restoredParagraphPropertiesXml: '<w:pPr/>',
			bookmarks: ['start'],
		},
		[schema.text('original')],
	);
	const doc = schema.node('doc', null, [paragraph]);
	view = new EditorView(document.body, {
		state: EditorState.create({ doc, selection: TextSelection.create(doc, 9) }),
	});
	expect(wordAddInSelection(() => view).writeText('a\nb\nc')).toBe(true);
	expect(view.state.doc.childCount).toBe(3);
	expect(view.state.doc.child(0).attrs).toEqual(paragraph.attrs);
	for (const index of [1, 2]) {
		const attrs = view.state.doc.child(index).attrs;
		expect(attrs).toMatchObject({ style: 'ListParagraph', numId: 3, ilvl: 0 });
		expect(attrs).toMatchObject({
			id: '',
			listLabelText: null,
			markRevision: null,
			formatRevision: null,
			sourceParagraphPropertiesXml: null,
			restoredParagraphPropertiesXml: null,
			bookmarks: [],
		});
	}
});

it('resolves the Word view on each request, so a pane opened before the view still works', async () => {
	let current: EditorView | undefined;
	const selection = await loadWordAddInSelection(() => current);
	expect(selection.readText()).toBeUndefined();
	expect(selection.writeText('early')).toBe(false);
	const doc = schema.node('doc', null, [schema.node('paragraph', null, [schema.text('later')])]);
	view = new EditorView(document.body, {
		state: EditorState.create({ doc, selection: TextSelection.create(doc, 1, 6) }),
	});
	current = view;
	expect(selection.readText()).toBe('later');
	view.destroy();
	expect(selection.readText()).toBeUndefined();
	view = undefined;
});
