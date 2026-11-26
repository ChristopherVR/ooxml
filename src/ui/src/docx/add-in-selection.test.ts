import { afterEach, expect, it } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { wordAddInSelection } from './add-in-selection';
import { schema } from './schema';

let view: EditorView | undefined;
afterEach(() => view?.destroy());

it('reads a read-only selection and writes multiline plain text as paragraphs when editable', () => {
	const doc = schema.node('doc', null, [schema.node('paragraph', null, [schema.text('original')])]);
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
});
