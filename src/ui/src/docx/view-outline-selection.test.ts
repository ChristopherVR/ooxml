// @vitest-environment jsdom
import { createDocument } from 'ooxml-core/docx';
import { indexOutline, OUTLINE_BODY, visibleOutline } from 'ooxml-core/docx/ui';
import { Schema } from 'prosemirror-model';
import { AllSelection, EditorState, TextSelection } from 'prosemirror-state';
import { DecorationSet } from 'prosemirror-view';
import { expect, it } from 'vitest';
import {
	deleteCollapsedOutlineBranches,
	outlineKey,
	outlineSelectionPlugin,
} from './view-outline-selection';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: {
			content: 'text*',
			attrs: { id: { default: null }, outlineLevel: { default: null } },
		},
		text: {},
	},
});

function setup() {
	const model = createDocument();
	const doc = schema.node('doc', null, [
		schema.node('paragraph', { id: 'a', outlineLevel: 1 }, schema.text('A')),
		schema.node('paragraph', { id: 'b' }, schema.text('B')),
		schema.node('paragraph', { id: 'c', outlineLevel: 1 }, schema.text('C')),
	]);
	const collapsed = new Set<string>();
	let level = OUTLINE_BODY;
	const plugin = outlineSelectionPlugin((doc) => {
		const entries = indexOutline(doc, model);
		return {
			entries,
			visible: visibleOutline(entries, level, collapsed),
			decorations: DecorationSet.empty,
			collapsed,
		};
	});
	return {
		state: EditorState.create({ doc, plugins: [plugin] }),
		collapsed,
		setLevel(next: number) {
			level = next;
		},
	};
}

it('moves text selections out of hidden blocks in either direction', () => {
	const { state, collapsed } = setup();
	collapsed.add('a');
	const projected = state.apply(state.tr.setMeta(outlineKey, true));
	const select = (state: EditorState, pos: number) =>
		state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));
	const forward = select(projected, 4);
	expect(forward.selection.$head.parent.attrs.id).toBe('c');
	expect(select(forward, 4).selection.$head.parent.attrs.id).toBe('a');

	const range = projected.apply(
		projected.tr.setSelection(TextSelection.create(projected.doc, 1, 4)),
	);
	expect(range.selection.anchor).toBe(1);
	expect(range.selection.$head.parent.attrs.id).toBe('c');
});

it('removes a fully selected collapsed heading but keeps the branch when only a caret is set', () => {
	const { state, collapsed } = setup();
	collapsed.add('a');
	const projected = state.apply(state.tr.setMeta(outlineKey, true));
	expect(deleteCollapsedOutlineBranches(projected)).toBeUndefined();
	const selected = projected.apply(
		projected.tr.setSelection(TextSelection.create(projected.doc, 1, 2)),
	);
	const deletion = deleteCollapsedOutlineBranches(selected);
	expect(deletion?.doc.textContent).toBe('C');
});

it('prevents editing when every block is filtered and restores a caret when shown', () => {
	const { state, setLevel } = setup();
	const body = state.apply(
		state.tr.setNodeAttribute(0, 'outlineLevel', null).setNodeAttribute(6, 'outlineLevel', null),
	);
	setLevel(1);
	const filtered = body.apply(body.tr.setMeta(outlineKey, true));
	expect(filtered.selection).toBeInstanceOf(AllSelection);
	expect(filtered.apply(filtered.tr.insertText('X')).doc.textContent).toBe('ABC');
	setLevel(OUTLINE_BODY);
	const shown = filtered.apply(filtered.tr.setMeta(outlineKey, true));
	expect(shown.selection).toBeInstanceOf(TextSelection);
	expect(shown.selection.$head.parent.attrs.id).toBe('a');
});
