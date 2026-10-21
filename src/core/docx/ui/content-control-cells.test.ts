import { describe, expect, it } from 'vitest';
import { Schema, type Node as ProseMirrorNode } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
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
		tableRow: {
			content: 'tableCell+',
			attrs: { properties: { default: null } },
			toDOM: () => ['tr', 0],
		},
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

	it('deletes a selection running from a text cell into the next as ordinary text', () => {
		const control: ContentControl = { id: '62', level: 'cell', kind: 'text' };
		const cell = schema.node('tableCell', { sourceCellKey: 'c1', contentControl: control }, [
			paragraph('x', { id: 'c1' }),
		]);
		const after = schema.node('tableCell', null, [paragraph('Next')]);
		const doc = schema.node('doc', null, [
			schema.node('table', null, [schema.node('tableRow', null, [cell, after])]),
		]);
		const state = EditorState.create({ schema, doc, plugins: [contentControlPlugin()] });
		const span = controlSpans(state.doc)[0]!;
		// From the start of the cell's text into the second cell's text.
		const into = state.apply(
			state.tr.setSelection(TextSelection.create(state.doc, span.from + 2, span.to + 4)),
		);
		expect(deleteToPlaceholder(into, 'backward')).toBeUndefined();
	});

	it('types over an inline placeholder at its edge inside a cell control', () => {
		const name: ContentControl = {
			id: '65',
			level: 'inline',
			kind: 'text',
			showingPlaceholder: true,
		};
		const prompt = schema.text('Enter a name.', [
			style('PlaceholderText'),
			schema.marks.contentControl!.create({ data: JSON.stringify(name) }),
		]);
		const content = [
			schema.node('paragraph', { id: 'c1' }, [schema.text('Name: ', [style('Normal')]), prompt]),
		];
		const state = stateOf({ id: '62', level: 'cell', kind: 'richText' }, content);
		const inline = controlSpans(state.doc).find((span) => span.level === 'inline')!;
		expect(controlAt(controlSpans(state.doc), inline.from, true)?.control.id).toBe('65');
		const typed = typeOverPlaceholder(state, inline.from, 'A');
		if (typed === 'blocked' || !typed) throw new Error('expected a transaction');
		expect(cellOf(state.apply(typed)).textContent).toBe('Name: A');
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

	it('leaves a cell alone when replacing its content would delete a control that may not be', () => {
		const fixed: ContentControl = {
			id: '64',
			level: 'inline',
			kind: 'richText',
			lock: 'sdtLocked',
		};
		const inner = schema.text('Fixed', [
			style('Normal'),
			schema.marks.contentControl!.create({ data: JSON.stringify(fixed) }),
		]);
		const content = [
			schema.node('paragraph', { id: 'c1' }, [schema.text('Note ', [style('Normal')]), inner]),
		];
		const text = stateOf({ id: '62', level: 'cell', kind: 'text' }, content);
		const span = controlSpans(text.doc)[0]!;
		// The whole text of the cell selected, from its first character in plain text.
		const all = text.apply(
			text.tr.setSelection(TextSelection.create(text.doc, span.from + 2, span.to - 2)),
		);
		expect(deleteToPlaceholder(all, 'backward')).toBeUndefined();
		// The ordinary deletion that runs instead is refused by the lock guard.
		expect(all.applyTransaction(all.tr.deleteSelection()).state.doc.eq(all.doc)).toBe(true);
		const list = stateOf(choice({ showingPlaceholder: false }), content);
		expect(chooseListItem(list, '61', 1)).toBeUndefined();
	});

	it('picks in the cell that was operated when two cell controls share an id', () => {
		const label = schema.node('tableCell', null, [paragraph('Status')]);
		const prompt = (id: string) => paragraph('Choose an item.', { id }, [style('PlaceholderText')]);
		const first = schema.node('tableCell', { sourceCellKey: 'c1', contentControl: choice() }, [
			prompt('c1'),
		]);
		const copy = schema.node('tableCell', { sourceCellKey: 'c2', contentControl: choice() }, [
			prompt('c2'),
		]);
		const doc = schema.node('doc', null, [
			schema.node('table', null, [schema.node('tableRow', null, [label, first, copy])]),
		]);
		const state = EditorState.create({ schema, doc, plugins: [contentControlPlugin()] });
		const second = controlSpans(state.doc)[1]!;
		const next = state.apply(chooseListItem(state, '61', 1, second.from + 3)!);
		const row = next.doc.firstChild!.firstChild!;
		expect([row.child(1).textContent, row.child(2).textContent]).toEqual([
			'Choose an item.',
			'Secondary',
		]);
	});

	it('gives a pasted copy of a cell or inline control a new id, and a moved one its own', () => {
		const name: ContentControl = { id: '65', level: 'inline', kind: 'text' };
		const named = (text: string, marks = [style('Normal')]) =>
			schema.text(text, [
				...marks,
				schema.marks.contentControl!.create({ data: JSON.stringify(name) }),
			]);
		const state = stateOf(choice(), [
			schema.node('paragraph', { id: 'c1' }, [
				named('Ann'),
				named('e', [schema.marks.bold!.create()]),
			]),
		]);
		const plugin = contentControlPlugin();
		const transform = plugin.props.transformPasted!;
		const slice = state.doc.slice(0, state.doc.content.size);
		const pasted = (view: Partial<EditorView>) =>
			controlSpans(
				schema.node('doc', null, transform.call(plugin, slice, view as EditorView, false).content),
			);
		const copies = pasted({ state, dragging: null });
		expect(copies.map((span) => span.level)).toEqual(['cell', 'inline']);
		expect(copies[0]!.control).toMatchObject({ kind: 'dropDownList', items: choice().items });
		for (const [index, id] of ['61', '65'].entries()) {
			expect(copies[index]!.control.id).not.toBe(id);
			expect(copies[index]!.control.id).toMatch(/^\d+$/);
		}
		// Both text nodes of the inline control keep one control: one span.
		expect(copies[1]!.to - copies[1]!.from).toBe(4);
		// Moving the control by dragging keeps its id, as does pasting where the id is not in use.
		const moved = pasted({ state, dragging: { slice, move: true } });
		expect(moved.map((span) => span.control.id)).toEqual(['61', '65']);
		const elsewhere = EditorState.create({
			schema,
			doc: schema.node('doc', null, [paragraph('x')]),
		});
		expect(pasted({ state: elsewhere, dragging: null }).map((span) => span.control.id)).toEqual([
			'61',
			'65',
		]);
	});

	it('removes a temporary row control once its row is edited, and not for an edit elsewhere', () => {
		const item: ContentControl = { id: '66', level: 'row', kind: 'richText', temporary: true };
		const row = (text: string, properties: unknown = null) =>
			schema.node('tableRow', { properties }, [schema.node('tableCell', null, [paragraph(text)])]);
		const doc = schema.node('doc', null, [
			schema.node('table', null, [
				row('Title'),
				row('Line', { heightTwips: 400, contentControl: item }),
			]),
		]);
		const state = EditorState.create({ schema, doc, plugins: [contentControlPlugin()] });
		const rowOf = (s: EditorState) => s.doc.firstChild!.child(1).attrs.properties;
		const elsewhere = state.apply(state.tr.insertText('!', 4));
		expect(rowOf(elsewhere)).toMatchObject({ contentControl: { id: '66' } });
		const edited = elsewhere.apply(elsewhere.tr.insertText('!', 18));
		expect(edited.doc.firstChild!.child(1).textContent).toBe('Li!ne');
		expect(rowOf(edited)).toEqual({ heightTwips: 400 });
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
