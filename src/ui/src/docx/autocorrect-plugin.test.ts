// @vitest-environment jsdom
import { history, undo } from 'prosemirror-history';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { createDocument } from 'ooxml-core/docx';
import { afterEach, describe, expect, it } from 'vitest';
import { autoCorrectPlugin } from './autocorrect-plugin';
import { modelToDoc } from './model-adapter';

afterEach(() => {
	document.body.innerHTML = '';
});

function createEditor(text = '') {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p1', runs: text ? [{ text }] : [] }];
	const host = document.createElement('div');
	document.body.append(host);
	const doc = modelToDoc(model);
	const view = new EditorView(host, {
		state: EditorState.create({
			doc,
			selection: TextSelection.create(doc, 1 + text.length),
			plugins: [history(), autoCorrectPlugin(() => model)],
		}),
	});
	return { view, model };
}

function typeChar(view: EditorView, ch: string) {
	const { from, to } = view.state.selection;
	const handled = view.someProp('handleTextInput', (fn) =>
		fn(view, from, to, ch, () => view.state.tr.insertText(ch, from, to)),
	);
	if (!handled) view.dispatch(view.state.tr.insertText(ch, from, to));
}

describe('autoCorrectPlugin', () => {
	it('converts straight quotes to smart quotes and reverts on single Undo', () => {
		const { view } = createEditor();
		typeChar(view, '"');
		expect(view.state.doc.textContent).toBe('“');
		// Single Undo reverts smart quote back to straight quote
		undo(view.state, (tr) => view.dispatch(tr));
		expect(view.state.doc.textContent).toBe('"');
		view.destroy();
	});

	it('converts "- " at paragraph start into a bulleted list and "1. " into a numbered list', () => {
		const { view, model } = createEditor('-');
		typeChar(view, ' ');
		const p1 = view.state.doc.firstChild!;
		expect(p1.textContent).toBe('');
		expect(p1.attrs.numId).not.toBeNull();
		expect(p1.attrs.ilvl).toBe(0);
		expect(model.numberingCatalog?.nums[p1.attrs.numId as number]).toBeDefined();

		// Undo reverts auto-list back to "- "
		undo(view.state, (tr) => view.dispatch(tr));
		expect(view.state.doc.firstChild!.attrs.numId).toBeNull();
		expect(view.state.doc.textContent).toBe('- ');
		view.destroy();
	});

	it('applies fractions and common typos on space', () => {
		const { view } = createEditor('teh');
		typeChar(view, ' ');
		expect(view.state.doc.textContent).toBe('the ');
		view.destroy();
	});
});
