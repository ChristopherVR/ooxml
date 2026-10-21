import { describe, expect, it, vi } from 'vitest';
import { Fragment, Schema, Slice, type Node as ProseMirrorNode } from 'prosemirror-model';
import { closeHistory, history, redo, undo } from 'prosemirror-history';
import { EditorState, TextSelection } from 'prosemirror-state';
import { ReplaceAroundStep } from 'prosemirror-transform';
import type { EditorView } from 'prosemirror-view';
import type { ContentControl } from '../model';
import { markSpecs } from './schema-marks';
import {
	chooseListItem,
	deleteToPlaceholder,
	newControlId,
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
		table: {
			group: 'block',
			content: 'tableRow+',
			attrs: { contentControl: { default: null } },
			toDOM: () => ['table', ['tbody', 0]],
		},
		tableRow: {
			content: 'tableCell+',
			attrs: { properties: { default: null } },
			toDOM: () => ['tr', 0],
		},
		tableCell: {
			content: '(paragraph | table)+',
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

	it('undoes and redoes a command on a cell whose control may not be deleted', () => {
		// Undo and redo are not the control's own edit, and they give the cell the control it had:
		// that deletes no control, and refusing them would leave every earlier edit undone too.
		const box: ContentControl = {
			id: '63',
			level: 'cell',
			kind: 'checkbox',
			checked: false,
			lock: 'sdtLocked',
		};
		const loaded = stateOf(box, [paragraph('\u2610', { id: 'c1' })]);
		let state = EditorState.create({
			schema,
			doc: loaded.doc,
			plugins: [history(), contentControlPlugin()],
		});
		const typed = state.apply(state.tr.insertText('!', state.doc.content.size - 3));
		state = typed.apply(closeHistory(toggleCheckboxControl(typed, '63')!));
		expect(cellOf(state).attrs.contentControl.checked).toBe(true);
		const run = (command: typeof undo) =>
			command(state, (tr) => {
				state = state.apply(tr);
			});
		expect(run(undo)).toBe(true);
		expect(cellOf(state).attrs.contentControl).toEqual(box);
		expect(cellOf(state).textContent).toBe('\u2610!');
		expect(run(undo)).toBe(true);
		expect(cellOf(state).textContent).toBe('\u2610');
		expect(run(redo)).toBe(true);
		expect(run(redo)).toBe(true);
		expect(cellOf(state).attrs.contentControl.checked).toBe(true);
		// The check box holds its glyph alone, as the toggle wrote it.
		expect(cellOf(state).textContent).toBe('\u2612');
	});

	it('applies a content lock of a cell control to the cell node, not to the table around it', () => {
		// The cell is what a cell control holds: other attributes for it edit that, and the table
		// around it takes them as before.
		const fixed: ContentControl = {
			id: '67',
			level: 'cell',
			kind: 'richText',
			lock: 'contentLocked',
		};
		const state = stateOf(fixed, [paragraph('Kept', { id: 'c1' })]);
		const at = controlSpans(state.doc)[0]!.from;
		const marked = (pos: number, attrs: Record<string, unknown>) =>
			violatesControlLock(state.tr.setNodeMarkup(pos, undefined, attrs), state);
		expect(marked(at, { ...state.doc.nodeAt(at)!.attrs, sourceCellKey: 'c9' })).toBe(true);
		expect(marked(at + 1, { ...state.doc.nodeAt(at + 1)!.attrs, id: 'c2' })).toBe(true);
		expect(marked(0, { contentControl: null })).toBe(false);
	});

	it('lets the node of a control that may not be deleted take other attributes', () => {
		// Table Properties, shading and paragraph formatting give a node other attributes in one
		// step over the whole node (`setNodeMarkup`), which keeps the node, what it holds and its
		// control: that deletes no control. Taking the control off that way still does.
		const fixed = (level: ContentControl['level']): ContentControl => ({
			id: '66',
			level,
			kind: 'richText',
			lock: 'sdtLocked',
		});
		const marked = (state: EditorState, at: number, attrs: Record<string, unknown>) =>
			violatesControlLock(state.tr.setNodeMarkup(at, undefined, attrs), state);
		const cell = stateOf(fixed('cell'), [paragraph('Kept', { id: 'c1' })]);
		const at = controlSpans(cell.doc)[0]!.from;
		const attrs = cell.doc.nodeAt(at)!.attrs;
		expect(marked(cell, at, { ...attrs, sourceCellKey: 'c9' })).toBe(false);
		expect(marked(cell, at, { ...attrs, contentControl: null })).toBe(true);
		// The table that holds it.
		expect(marked(cell, 0, { contentControl: null })).toBe(false);
		// A row and a paragraph the same way.
		const properties = { contentControl: fixed('row') };
		const row = EditorState.create({
			schema,
			plugins: [contentControlPlugin()],
			doc: schema.node('doc', null, [
				schema.node('table', null, [
					schema.node('tableRow', { properties }, [
						schema.node('tableCell', null, [paragraph('Row', { id: 'r1' })]),
					]),
				]),
			]),
		});
		expect(marked(row, 1, { properties: { ...properties, heightTwips: 400 } })).toBe(false);
		expect(marked(row, 1, { properties: null })).toBe(true);
		const block = EditorState.create({
			schema,
			plugins: [contentControlPlugin()],
			doc: schema.node('doc', null, [
				paragraph('Block', { id: 'b1', contentControl: fixed('block') }),
				paragraph('After'),
			]),
		});
		expect(marked(block, 0, { id: 'b2', contentControl: fixed('block') })).toBe(false);
		expect(marked(block, 0, { id: 'b1', contentControl: null })).toBe(true);
		// A step of that shape whose new node brings content of its own, or one over text, can
		// delete a control: it is no such change.
		const own = block.doc.firstChild!;
		const added = new ReplaceAroundStep(
			0,
			own.nodeSize,
			1,
			own.nodeSize - 1,
			new Slice(Fragment.from(own.type.create(own.attrs, schema.text('!'))), 0, 0),
			1,
			true,
		);
		expect(violatesControlLock(block.tr.step(added), block)).toBe(true);
		const run = schema.text('ab', [
			schema.marks.contentControl!.create({ data: JSON.stringify(fixed('inline')) }),
		]);
		const inline = EditorState.create({
			schema,
			plugins: [contentControlPlugin()],
			doc: schema.node('doc', null, [schema.node('paragraph', null, [run])]),
		});
		const over = new ReplaceAroundStep(
			1,
			3,
			2,
			2,
			new Slice(Fragment.from(schema.text('x')), 0, 0),
			1,
		);
		expect(violatesControlLock(inline.tr.step(over), inline)).toBe(true);
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

	it('types as usual over a selection running from a placeholder cell into the next', () => {
		const control: ContentControl = {
			id: '62',
			level: 'cell',
			kind: 'text',
			showingPlaceholder: true,
		};
		const cell = schema.node('tableCell', { sourceCellKey: 'c1', contentControl: control }, [
			placeholder('Enter a number.'),
		]);
		const after = schema.node('tableCell', null, [paragraph('Next')]);
		const doc = schema.node('doc', null, [
			schema.node('table', null, [schema.node('tableRow', null, [cell, after])]),
		]);
		const state = EditorState.create({ schema, doc, plugins: [contentControlPlugin()] });
		const span = controlSpans(state.doc)[0]!;
		// From inside the placeholder into the second cell's text: not the placeholder's to replace.
		expect(typeOverPlaceholder(state, span.from + 4, '7', span.to + 4)).toBeUndefined();
		// Within the cell, typing replaces the whole placeholder.
		const within = typeOverPlaceholder(state, span.from + 4, '7', span.to - 3);
		if (within === 'blocked' || !within) throw new Error('expected a transaction');
		const row = state.apply(within).doc.firstChild!.firstChild!;
		expect([row.child(0).textContent, row.child(1).textContent]).toEqual(['7', 'Next']);
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

	it('leaves a cell alone when replacing its content would delete a table or row control that may not be', () => {
		const fixed = (level: 'row' | 'block'): ContentControl => ({
			id: '65',
			level,
			kind: 'richText',
			lock: 'sdtLocked',
		});
		for (const level of ['row', 'block'] as const) {
			const nested = schema.node(
				'table',
				level === 'block' ? { contentControl: fixed(level) } : null,
				[
					schema.node(
						'tableRow',
						level === 'row' ? { properties: { contentControl: fixed(level) } } : null,
						[schema.node('tableCell', null, [paragraph('Inner', { id: 'n1' })])],
					),
				],
			);
			const text = stateOf({ id: '62', level: 'cell', kind: 'richText' }, [
				paragraph('Before', { id: 'c1' }),
				nested,
				paragraph('After', { id: 'c2' }),
			]);
			const span = controlSpans(text.doc)[0]!;
			const all = text.apply(
				text.tr.setSelection(TextSelection.create(text.doc, span.from + 2, span.to - 2)),
			);
			expect(deleteToPlaceholder(all, 'backward')).toBeUndefined();
			// The ordinary deletion that runs instead is refused by the lock guard.
			expect(all.applyTransaction(all.tr.deleteSelection()).state.doc.eq(all.doc)).toBe(true);
		}
	});

	it('takes a control around the replaced paragraph away with its text', () => {
		// The paragraph's own control is deleted with the content a pick or a toggle replaces, as a
		// control that may be deleted (here one whose content may not be edited) can be.
		const notes: ContentControl = {
			id: '68',
			level: 'block',
			kind: 'richText',
			lock: 'contentLocked',
		};
		const content = [paragraph('Fixed text', { id: 'c1', contentControl: notes })];
		const list = stateOf(choice({ showingPlaceholder: false }), content);
		const picked = cellOf(list.apply(chooseListItem(list, '61', 1)!));
		expect(picked.textContent).toBe('Secondary');
		expect(picked.firstChild!.attrs).toMatchObject({ id: 'c1', contentControl: null });
		const box = stateOf({ id: '63', level: 'cell', kind: 'checkbox', checked: false }, content);
		const toggled = cellOf(box.apply(toggleCheckboxControl(box, '63')!));
		expect(toggled.textContent).toBe('☒');
		expect(toggled.firstChild!.attrs.contentControl).toBeNull();
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

	it('removes a temporary control around two cells, or around a table in a block, as one control', () => {
		const pair = choice({ kind: 'richText', temporary: true, showingPlaceholder: false });
		const cell = (text: string, control: ContentControl | null) =>
			schema.node('tableCell', { contentControl: control }, [paragraph(text)]);
		const twoCells = (separate: boolean) =>
			EditorState.create({
				schema,
				plugins: [contentControlPlugin()],
				doc: schema.node('doc', null, [
					schema.node('table', null, [
						schema.node('tableRow', null, [
							cell('One', pair),
							...(separate ? [cell('Gap', null)] : []),
							cell('Two', pair),
						]),
					]),
				]),
			});
		const controls = (state: EditorState) => controlSpans(state.doc).length;
		// One control around two cells: an edit in either takes it off both.
		const joined = twoCells(false);
		expect(controls(joined.apply(joined.tr.insertText('!', 5)))).toBe(0);
		// Two controls that share an id: the one not edited stays.
		const separate = twoCells(true);
		expect(controls(separate.apply(separate.tr.insertText('!', 5)))).toBe(1);
		// A block control holding a paragraph, a table and a paragraph is one control.
		const block = { ...pair, level: 'block' as const };
		const doc = schema.node('doc', null, [
			paragraph('Before', { contentControl: block }),
			schema.node('table', null, [schema.node('tableRow', null, [cell('Inside', null)])]),
			paragraph('After', { contentControl: block }),
		]);
		const state = EditorState.create({ schema, doc, plugins: [contentControlPlugin()] });
		const blocks = (s: EditorState) => {
			const controls: unknown[] = [];
			s.doc.forEach((node) => controls.push(node.attrs.contentControl?.id ?? null));
			return controls;
		};
		const edited = state.apply(state.tr.insertText('!', doc.content.size - 2));
		expect(blocks(edited)).toEqual([null, null, null]);
		// The table is part of its content: editing only the table removes it too.
		const inTable = state.apply(state.tr.insertText('!', 14));
		expect(inTable.doc.child(1).textContent).toBe('In!side');
		expect(blocks(inTable)).toEqual([null, null, null]);
		// One holding a paragraph of a nested block control between two of its own.
		const nested = { id: '67', level: 'block' as const, kind: 'richText' as const };
		const around = schema.node('doc', null, [
			paragraph('First', { contentControl: block }),
			paragraph('Inner', { contentControl: nested }),
			paragraph('Last', { contentControl: block }),
		]);
		const outer = EditorState.create({ schema, doc: around, plugins: [contentControlPlugin()] });
		const ids: unknown[] = [];
		outer
			.apply(outer.tr.insertText('!', 2))
			.doc.forEach((node) => ids.push(node.attrs.contentControl?.id ?? null));
		expect(ids).toEqual([null, '67', null]);
	});

	it('removes a temporary block control from a table it holds once the table is edited', () => {
		const block: ContentControl = { id: '69', level: 'block', kind: 'richText', temporary: true };
		const table = (control: ContentControl | null) =>
			schema.node('table', { contentControl: control }, [
				schema.node('tableRow', null, [schema.node('tableCell', null, [paragraph('Inside')])]),
			]);
		const controls = (state: EditorState) => {
			const ids: unknown[] = [];
			state.doc.forEach((node) => ids.push(node.attrs.contentControl?.id ?? null));
			return ids;
		};
		// Around the table alone, and around a paragraph, the table and a paragraph.
		for (const blocks of [
			[table(block), paragraph('Elsewhere')],
			[
				paragraph('Before', { contentControl: block }),
				table(block),
				paragraph('After', { contentControl: block }),
				paragraph('Elsewhere'),
			],
		]) {
			const state = EditorState.create({
				schema,
				doc: schema.node('doc', null, blocks),
				plugins: [contentControlPlugin()],
			});
			const outside = state.apply(state.tr.insertText('!', state.doc.content.size - 1));
			expect(controls(outside)).toEqual(controls(state));
			let at = -1;
			state.doc.descendants((node, pos) => {
				if (at < 0 && node.isText && node.text === 'Inside') at = pos + 2;
			});
			const edited = state.apply(state.tr.insertText('!', at));
			expect(controls(edited).every((id) => id === null)).toBe(true);
		}
	});

	it('keeps the formatting of deleted cell text for what is typed next when a table control goes', () => {
		const block: ContentControl = { id: '69', level: 'block', kind: 'richText', temporary: true };
		const bold = [style('Normal'), schema.marks.bold!.create()];
		const doc = schema.node('doc', null, [
			schema.node('table', { contentControl: block }, [
				schema.node('tableRow', null, [
					schema.node('tableCell', null, [paragraph('Inside', {}, bold)]),
				]),
			]),
		]);
		const state = EditorState.create({ schema, doc, plugins: [contentControlPlugin()] });
		// The whole text of the cell (table, row, cell and paragraph open first).
		const all = state.apply(state.tr.setSelection(TextSelection.create(doc, 4, 10)));
		const deleted = all.apply(all.tr.deleteSelection());
		expect(deleted.doc.firstChild!.attrs.contentControl).toBeNull();
		const typed = deleted.apply(deleted.tr.insertText('Z'));
		expect(typed.doc.textContent).toBe('Z');
		let marks: string[] = [];
		typed.doc.descendants((node) => {
			if (node.isText) marks = node.marks.map((mark) => mark.type.name);
		});
		expect(marks).toContain('bold');
	});

	it('takes a temporary control off only itself when another one shares its id', () => {
		// Files from other tools repeat ids: here a cell control and a block control in its cell.
		const cellControl = (temporary: boolean): ContentControl => ({
			id: '5',
			level: 'cell',
			kind: 'richText',
			...(temporary ? { temporary } : {}),
		});
		const blockControl = (temporary: boolean): ContentControl => ({
			...cellControl(temporary),
			level: 'block',
		});
		for (const temporaryCell of [true, false]) {
			const state = stateOf(cellControl(temporaryCell), [
				paragraph('Note', { id: 'c1', contentControl: blockControl(!temporaryCell) }),
			]);
			const edited = cellOf(state.apply(state.tr.insertText('!', 16)));
			expect(edited.textContent).toBe('No!te');
			expect([
				edited.attrs.contentControl?.id,
				edited.firstChild!.attrs.contentControl?.id,
			]).toEqual(temporaryCell ? [undefined, '5'] : ['5', undefined]);
		}
	});

	it('gives a new control an id that no row control has', () => {
		const item: ContentControl = { id: '1000', level: 'row', kind: 'richText' };
		const doc = schema.node('doc', null, [
			schema.node('table', null, [
				schema.node('tableRow', { properties: { contentControl: item } }, [
					schema.node('tableCell', null, [paragraph('Line')]),
				]),
			]),
		]);
		// The first id drawn is the row control's, the second one is free.
		const draw = (id: number) => (id - 0.5) / (2 ** 31 - 1);
		const random = vi.spyOn(Math, 'random');
		try {
			random.mockReturnValueOnce(draw(1000)).mockReturnValueOnce(draw(1001));
			expect(newControlId(doc)).toBe('1001');
		} finally {
			random.mockRestore();
		}
	});

	it('gives a new or pasted control an id that no table control has', () => {
		const sheet: ContentControl = { id: '1000', level: 'block', kind: 'richText' };
		const doc = schema.node('doc', null, [
			schema.node('table', { contentControl: sheet }, [
				schema.node('tableRow', null, [schema.node('tableCell', null, [paragraph('Line')])]),
			]),
			paragraph('After'),
		]);
		const state = EditorState.create({ schema, doc, plugins: [contentControlPlugin()] });
		const draw = (id: number) => (id - 0.5) / (2 ** 31 - 1);
		const random = vi.spyOn(Math, 'random');
		try {
			random.mockReturnValueOnce(draw(1000)).mockReturnValueOnce(draw(1001));
			expect(newControlId(doc)).toBe('1001');
			// A copy of an inline control with the table control's id, pasted after it.
			const name: ContentControl = { id: '1000', level: 'inline', kind: 'text' };
			const copied = schema.node('doc', null, [
				schema.node('paragraph', null, [
					schema.text('Ann', [schema.marks.contentControl!.create({ data: JSON.stringify(name) })]),
				]),
			]);
			const plugin = contentControlPlugin();
			random.mockReturnValueOnce(draw(1002));
			const pasted = plugin.props.transformPasted!.call(
				plugin,
				copied.slice(0, copied.content.size),
				{ state, dragging: null } as unknown as EditorView,
				false,
			);
			const ids = controlSpans(schema.node('doc', null, pasted.content)).map(
				(span) => span.control.id,
			);
			expect(ids).toEqual(['1002']);
		} finally {
			random.mockRestore();
		}
	});

	it('leaves a control alone when a control around it forbids editing its content', () => {
		// The commands replace a control's content past the lock guard, so they look at the
		// controls it sits in: an outer cell or block control whose content is locked.
		const locked = (level: 'cell' | 'block'): ContentControl => ({
			id: '71',
			level,
			kind: 'richText',
			lock: 'contentLocked',
		});
		const list = choice({ showingPlaceholder: false });
		const inner = schema.node('table', null, [
			schema.node('tableRow', null, [
				schema.node('tableCell', { contentControl: list }, [paragraph('Primary', { id: 'n1' })]),
			]),
		]);
		const outer = schema.node('tableCell', { contentControl: locked('cell') }, [
			paragraph('Sources', { id: 'o1' }),
			inner,
		]);
		const nested = EditorState.create({
			schema,
			plugins: [contentControlPlugin()],
			doc: schema.node('doc', null, [
				schema.node('table', null, [schema.node('tableRow', null, [outer])]),
			]),
		});
		expect(chooseListItem(nested, '61', 1)).toBeUndefined();
		const pick: ContentControl = { ...list, id: '72', level: 'inline' };
		const box: ContentControl = { id: '73', level: 'inline', kind: 'checkbox', checked: false };
		const prompt: ContentControl = {
			id: '74',
			level: 'inline',
			kind: 'text',
			showingPlaceholder: true,
		};
		const run = (text: string, control: ContentControl) =>
			schema.text(text, [schema.marks.contentControl!.create({ data: JSON.stringify(control) })]);
		const block = EditorState.create({
			schema,
			plugins: [contentControlPlugin()],
			doc: schema.node('doc', null, [
				schema.node('paragraph', { contentControl: locked('block') }, [
					run('Primary', pick),
					run('\u2610', box),
					run('Enter a name.', prompt),
				]),
			]),
		});
		expect(chooseListItem(block, '72', 1)).toBeUndefined();
		expect(toggleCheckboxControl(block, '73')).toBeUndefined();
		const inPrompt = controlSpans(block.doc).find((span) => span.control.id === '74')!.from + 2;
		expect(typeOverPlaceholder(block, inPrompt, 'A')).toBe('blocked');
		// Not locked around, each of them is operated.
		const open = EditorState.create({
			schema,
			doc: schema.node('doc', null, [
				schema.node('paragraph', null, [run('Primary', pick), run('\u2610', box)]),
			]),
		});
		expect(chooseListItem(open, '72', 1)).toBeDefined();
		expect(toggleCheckboxControl(open, '73')).toBeDefined();
		// In a group control the controls follow their own locks, whatever the group's lock says.
		const group = EditorState.create({
			schema,
			doc: schema.node('doc', null, [
				schema.node('paragraph', { contentControl: { ...locked('block'), kind: 'group' } }, [
					run('Primary', pick),
					run('\u2610', box),
				]),
			]),
		});
		expect(chooseListItem(group, '72', 1)).toBeDefined();
		expect(toggleCheckboxControl(group, '73')).toBeDefined();
		// A row control locks its row's content too, for the commands and for typing.
		const rowLocked = EditorState.create({
			schema,
			plugins: [contentControlPlugin()],
			doc: schema.node('doc', null, [
				schema.node('table', null, [
					schema.node(
						'tableRow',
						{ properties: { contentControl: { ...locked('cell'), level: 'row' } } },
						[
							schema.node('tableCell', { contentControl: list }, [
								paragraph('Primary', { id: 'r1' }),
							]),
						],
					),
				]),
				paragraph('After'),
			]),
		});
		expect(chooseListItem(rowLocked, '61', 1)).toBeUndefined();
		expect(violatesControlLock(rowLocked.tr.insertText('!', 5), rowLocked)).toBe(true);
		const after = rowLocked.doc.content.size - 1;
		expect(violatesControlLock(rowLocked.tr.insertText('!', after), rowLocked)).toBe(false);
	});

	it('applies the lock of a block control around a table to the table', () => {
		// The editor keeps a block control around a whole table on the table node: its content is
		// the table, whose cells hold controls the commands operate past the lock guard.
		const list = choice({ showingPlaceholder: false });
		const around = (lock: ContentControl['lock']) =>
			EditorState.create({
				schema,
				plugins: [contentControlPlugin()],
				doc: schema.node('doc', null, [
					schema.node(
						'table',
						{ contentControl: { id: '75', level: 'block', kind: 'richText', lock } },
						[
							schema.node('tableRow', null, [
								schema.node('tableCell', null, [paragraph('Status', { id: 't1' })]),
								schema.node('tableCell', { contentControl: list }, [
									paragraph('Primary', { id: 't2' }),
								]),
							]),
						],
					),
					paragraph('After'),
				]),
			});
		const locked = around('contentLocked');
		const table = locked.doc.firstChild!.nodeSize;
		const after = locked.doc.content.size - 1;
		expect(chooseListItem(locked, '61', 1)).toBeUndefined();
		expect(violatesControlLock(locked.tr.insertText('!', 4), locked)).toBe(true);
		expect(violatesControlLock(locked.tr.insertText('!', after), locked)).toBe(false);
		// Its content cannot be edited, but the control can still be deleted whole.
		expect(violatesControlLock(locked.tr.delete(0, table), locked)).toBe(false);
		// One that may not be deleted keeps its content editable.
		const fixed = around('sdtLocked');
		expect(chooseListItem(fixed, '61', 1)).toBeDefined();
		expect(violatesControlLock(fixed.tr.insertText('!', 4), fixed)).toBe(false);
		expect(violatesControlLock(fixed.tr.delete(0, table), fixed)).toBe(true);
		// Other properties for the table (Table Properties) edit what the control holds: the one
		// whose content is locked refuses them, the one that may not be deleted takes them.
		const properties = (state: EditorState) =>
			state.tr.setNodeMarkup(0, undefined, state.doc.firstChild!.attrs);
		expect(violatesControlLock(properties(locked), locked)).toBe(true);
		expect(violatesControlLock(properties(fixed), fixed)).toBe(false);
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
