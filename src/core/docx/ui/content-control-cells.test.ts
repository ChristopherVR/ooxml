import { describe, expect, it } from 'vitest';
import { Schema, type Node as ProseMirrorNode } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { ContentControl } from '../model';
import { markSpecs } from './schema-marks';
import {
	chooseListItem,
	deleteToPlaceholder,
	toggleCheckboxControl,
	typeOverPlaceholder,
	updateControlProperties,
} from './content-control-commands';
import { contentControlPlugin, violatesControlLock } from './content-control-guard';
import { controlAt, controlSpans } from './content-control-spans';

const schema = new Schema({
	nodes: {
		doc: { content: 'block+' },
		paragraph: {
			group: 'block',
			content: 'inline*',
			attrs: { id: { default: null }, contentControl: { default: null } },
			toDOM: () => ['p', 0],
		},
		table: { group: 'block', content: 'tableRow+', toDOM: () => ['table', ['tbody', 0]] },
		tableRow: { content: 'tableCell+', toDOM: () => ['tr', 0] },
		tableCell: {
			content: 'paragraph+',
			attrs: { sourceCellKey: { default: '' }, contentControl: { default: null } },
			toDOM: () => ['td', 0],
		},
		text: { group: 'inline' },
	},
	marks: markSpecs,
});
const style = (id: string) => schema.marks.characterStyle!.create({ id });
const paragraph = (text: string, attrs: Record<string, unknown> = {}, marks = [style('Normal')]) =>
	schema.node('paragraph', attrs, text ? [schema.text(text, marks)] : []);
/** A one-row table: a label cell, then a cell holding `control` and `content`. */
const stateOf = (control: ContentControl, content: ProseMirrorNode[]) => {
	const label = schema.node('tableCell', null, [paragraph('Status')]);
	const cell = schema.node('tableCell', { sourceCellKey: 'c1', contentControl: control }, content);
	const doc = schema.node('doc', null, [
		schema.node('table', null, [schema.node('tableRow', null, [label, cell])]),
	]);
	return EditorState.create({ schema, doc, plugins: [contentControlPlugin()] });
};
const choice = (extra: Partial<ContentControl> = {}): ContentControl => ({
	id: '61',
	level: 'cell',
	kind: 'dropDownList',
	showingPlaceholder: true,
	items: [
		{ display: 'Primary', value: 'primary' },
		{ display: 'Secondary', value: 'secondary' },
	],
	...extra,
});
const placeholder = (text: string) => paragraph(text, { id: 'c1' }, [style('PlaceholderText')]);
const cellOf = (state: EditorState) => state.doc.firstChild!.firstChild!.child(1);
const stylesIn = (node: ProseMirrorNode) => {
	const ids: string[] = [];
	node.descendants((child) => {
		for (const mark of child.marks)
			if (mark.type.name === 'characterStyle') ids.push(mark.attrs.id);
	});
	return ids;
};

describe('content controls around table cells in the editor', () => {
	it('finds a cell control, and the control of a paragraph inside it first', () => {
		const state = stateOf(choice(), [placeholder('Choose an item.')]);
		const [span] = controlSpans(state.doc);
		expect(span).toMatchObject({ level: 'cell', control: { id: '61' } });
		expect(controlAt(controlSpans(state.doc), span!.from + 3)?.level).toBe('cell');
		const inner = stateOf(choice(), [
			paragraph('Note', { contentControl: { id: '62', level: 'block', kind: 'richText' } }),
		]);
		expect(
			controlAt(controlSpans(inner.doc), controlSpans(inner.doc)[0]!.from + 3)?.control.id,
		).toBe('62');
	});

	it('puts a picked entry in the cell and clears its placeholder', () => {
		const state = stateOf(choice(), [placeholder('Choose an item.'), paragraph('second line')]);
		const next = state.apply(chooseListItem(state, '61', 1)!);
		const cell = cellOf(next);
		expect(cell.childCount).toBe(1);
		expect(cell.textContent).toBe('Secondary');
		expect(cell.firstChild!.attrs.id).toBe('c1');
		expect(cell.attrs.contentControl.showingPlaceholder).toBeUndefined();
		expect(stylesIn(cell)).toEqual([]);
		expect(typeOverPlaceholder(next, controlSpans(next.doc)[0]!.from + 3, 'x')).toBe('blocked');
	});

	it('toggles a check box cell and leaves a locked one alone', () => {
		const box: ContentControl = { id: '63', level: 'cell', kind: 'checkbox', checked: false };
		const state = stateOf(box, [paragraph('☐', { id: 'c1' })]);
		const next = state.apply(toggleCheckboxControl(state, '63')!);
		expect(cellOf(next).textContent).toBe('☒');
		expect(cellOf(next).attrs.contentControl.checked).toBe(true);
		const locked = stateOf({ ...box, lock: 'contentLocked' }, [paragraph('☐', { id: 'c1' })]);
		expect(toggleCheckboxControl(locked, '63')).toBeUndefined();
		const typed = locked.tr.insertText('x', controlSpans(locked.doc)[0]!.from + 2);
		expect(violatesControlLock(typed, locked)).toBe(true);
	});

	it('replaces a text cell placeholder with what is typed, and brings it back on delete', () => {
		const { items: _items, ...list } = choice({ placeholderText: 'Enter a number.' });
		const text: ContentControl = { ...list, kind: 'text' };
		const state = stateOf(text, [placeholder('Enter a number.')]);
		const from = controlSpans(state.doc)[0]!.from;
		const typed = typeOverPlaceholder(state, from + 5, '4');
		if (typed === 'blocked' || !typed) throw new Error('expected a transaction');
		const next = state.apply(typed);
		expect(cellOf(next).textContent).toBe('4');
		expect(next.selection.from).toBe(from + 3);
		expect(cellOf(next).attrs.contentControl.showingPlaceholder).toBeUndefined();
		const cleared = next.apply(deleteToPlaceholder(next, 'backward')!);
		expect(cellOf(cleared).textContent).toBe('Enter a number.');
		expect(cellOf(cleared).attrs.contentControl.showingPlaceholder).toBe(true);
		expect(stylesIn(cellOf(cleared))).toEqual(['PlaceholderText']);
	});

	it('keeps the caret in a text cell whose placeholder comes back, not in the next cell', () => {
		const control: ContentControl = { id: '62', level: 'cell', kind: 'text' };
		const cell = schema.node('tableCell', { sourceCellKey: 'c1', contentControl: control }, [
			paragraph('x', { id: 'c1' }),
		]);
		const after = schema.node('tableCell', null, [paragraph('Next')]);
		const doc = schema.node('doc', null, [
			schema.node('table', null, [schema.node('tableRow', null, [cell, after])]),
		]);
		const state = EditorState.create({ schema, doc, plugins: [contentControlPlugin()] });
		// The caret after the cell's only character (table, row, cell and paragraph open first).
		const atEnd = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 5)));
		const restored = atEnd.apply(deleteToPlaceholder(atEnd, 'backward')!);
		expect(restored.doc.firstChild!.firstChild!.firstChild!.textContent).toBe(
			'Click or tap here to enter text.',
		);
		expect(restored.selection.$from.index(2)).toBe(0);
	});

	it('removes a temporary cell control with its first edit', () => {
		const state = stateOf(choice({ temporary: true }), [placeholder('Choose an item.')]);
		const next = state.apply(chooseListItem(state, '61', 0)!);
		expect(cellOf(next).textContent).toBe('Primary');
		expect(cellOf(next).attrs.contentControl).toBeNull();
		const text = stateOf(choice({ kind: 'text', temporary: true, showingPlaceholder: false }), [
			paragraph('Draft', { id: 'c1' }),
		]);
		const edited = text.apply(text.tr.insertText('s', controlSpans(text.doc)[0]!.from + 7));
		expect(cellOf(edited).attrs.contentControl).toBeNull();
	});

	it('stores new properties on the cell', () => {
		const state = stateOf(choice(), [placeholder('Choose an item.')]);
		const span = controlSpans(state.doc)[0]!;
		const next = state.apply(
			updateControlProperties(state, span, { ...span.control, alias: 'Kind' }),
		);
		expect(cellOf(next).attrs.contentControl.alias).toBe('Kind');
		expect(next.doc.textContent).toBe(state.doc.textContent);
		const caret = next.apply(next.tr.setSelection(TextSelection.create(next.doc, span.from + 3)));
		expect(controlAt(controlSpans(caret.doc), caret.selection.from)?.control.alias).toBe('Kind');
	});
});
