import { describe, expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { ContentControl, TextRun } from '../model';
import { markSpecs } from './schema-marks';
import { appendInlineNode, runToInlineNodes } from './run-adapter';
import {
	chooseDate,
	chooseListItem,
	deleteToPlaceholder,
	insertContentControl,
	toggleCheckboxControl,
	typeOverPlaceholder,
} from './content-control-commands';
import { contentControlPlugin, violatesControlLock } from './content-control-guard';
import { controlAt, controlSpans } from './content-control-spans';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: {
			content: 'inline*',
			attrs: { contentControl: { default: null } },
			toDOM: () => ['p', 0],
		},
		text: { group: 'inline' },
	},
	marks: markSpecs,
});
const control = (extra: Partial<ContentControl> = {}): ContentControl => ({
	id: '11',
	level: 'inline',
	kind: 'richText',
	...extra,
});
const stateOf = (runs: TextRun[], selection = 0) => {
	const nodes = runs.flatMap((run) => runToInlineNodes(run, schema));
	const doc = schema.node('doc', null, [schema.node('paragraph', null, nodes)]);
	const state = EditorState.create({ schema, doc, plugins: [contentControlPlugin()] });
	return selection
		? state.apply(state.tr.setSelection(TextSelection.create(doc, selection)))
		: state;
};
const text = (state: EditorState) => state.doc.textContent;

describe('content control marks round trip through the editor', () => {
	it('keeps the control on its runs and merges only runs of the same control', () => {
		const a = control({ alias: 'A' });
		const runs: TextRun[] = [
			{ text: 'one ', contentControl: a },
			{ text: 'two', contentControl: a, bold: true },
			{ text: 'three', contentControl: control({ id: '12' }) },
		];
		const state = stateOf(runs);
		const out: TextRun[] = [];
		state.doc.firstChild!.forEach((child) => appendInlineNode(out, child));
		expect(out.map((run) => [run.text, run.contentControl?.id])).toEqual([
			['one ', '11'],
			['two', '11'],
			['three', '12'],
		]);
		expect(out[0]?.contentControl?.alias).toBe('A');
		expect(controlSpans(state.doc).map((span) => span.control.id)).toEqual(['11', '12']);
	});
});

describe('content control commands', () => {
	it('toggles a checkbox with its glyph and font', () => {
		const state = stateOf([
			{
				text: '☐',
				contentControl: control({ kind: 'checkbox', checked: false }),
			},
		]);
		const next = state.apply(toggleCheckboxControl(state, '11')!);
		expect(text(next)).toBe('☒');
		const span = controlSpans(next.doc)[0]!;
		expect(span.control.checked).toBe(true);
		const font = next.doc.nodeAt(span.from)!.marks.find((mark) => mark.type.name === 'font');
		expect(font?.attrs.family).toBe('MS Gothic');
	});

	it('chooses a list item and a date', () => {
		const list = control({
			kind: 'dropDownList',
			showingPlaceholder: true,
			items: [
				{ display: 'Red', value: 'r' },
				{ display: 'Blue', value: 'b' },
			],
		});
		let state = stateOf([
			{ text: 'Choose an item.', contentControl: list, style: 'PlaceholderText' },
		]);
		state = state.apply(chooseListItem(state, '11', 1)!);
		expect(text(state)).toBe('Blue');
		expect(controlSpans(state.doc)[0]?.control.showingPlaceholder).toBeUndefined();
		expect(state.doc.nodeAt(1)!.marks.some((mark) => mark.type.name === 'characterStyle')).toBe(
			false,
		);
		let dated = stateOf([
			{ text: 'x', contentControl: control({ kind: 'date', dateFormat: 'yyyy-MM-dd' }) },
		]);
		dated = dated.apply(chooseDate(dated, '11', new Date(2026, 0, 9))!);
		expect(text(dated)).toBe('2026-01-09');
		expect(controlSpans(dated.doc)[0]?.control.fullDate).toBe('2026-01-09T00:00:00Z');
	});

	it('typing over a placeholder replaces it; a drop-down takes no typing', () => {
		const placeholder = control({ kind: 'text', showingPlaceholder: true });
		const state = stateOf([
			{ text: 'Before ' },
			{ text: 'Click here', contentControl: placeholder, style: 'PlaceholderText' },
		]);
		const tr = typeOverPlaceholder(state, 10, 'H');
		if (tr === 'blocked' || !tr) throw new Error('expected a transaction');
		const next = state.apply(tr);
		expect(text(next)).toBe('Before H');
		expect(controlSpans(next.doc)[0]?.control.showingPlaceholder).toBeUndefined();
		const drop = stateOf([{ text: 'Red', contentControl: control({ kind: 'dropDownList' }) }]);
		expect(typeOverPlaceholder(drop, 2, 'x')).toBe('blocked');
	});

	it('deleting all text brings the placeholder back', () => {
		const state = stateOf([{ text: 'Hi', contentControl: control({ kind: 'text' }) }], 1);
		const all = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 3)));
		const next = all.apply(deleteToPlaceholder(all, 'backward')!);
		expect(text(next)).toBe('Click or tap here to enter text.');
		expect(controlSpans(next.doc)[0]?.control.showingPlaceholder).toBe(true);
	});

	it('keeps a control whose own placeholder text is empty when its text is deleted', () => {
		const empty = control({ kind: 'text', placeholderText: '' });
		const state = stateOf([{ text: 'DOC-1', contentControl: empty }]);
		const all = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 6)));
		const next = all.apply(deleteToPlaceholder(all, 'backward')!);
		expect(text(next)).toBe('Click or tap here to enter text.');
		expect(controlSpans(next.doc)).toHaveLength(1);
	});

	it('brings back the control’s own placeholder, in the Placeholder Text style', () => {
		const own = control({ kind: 'text', placeholderText: 'Enter the document number.' });
		const state = stateOf([{ text: 'DOC-1', contentControl: own, style: 'Strong' }]);
		const all = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 6)));
		const next = all.apply(deleteToPlaceholder(all, 'backward')!);
		expect(text(next)).toBe('Enter the document number.');
		const styles = next.doc
			.nodeAt(1)!
			.marks.filter((mark) => mark.type.name === 'characterStyle')
			.map((mark) => mark.attrs.id);
		expect(styles).toEqual(['PlaceholderText']);
		const typed = typeOverPlaceholder(next, 3, 'D');
		if (typed === 'blocked' || !typed) throw new Error('expected a transaction');
		const value = next.apply(typed);
		expect(value.doc.nodeAt(1)!.marks.some((mark) => mark.attrs.id === 'PlaceholderText')).toBe(
			false,
		);
	});

	it('inserts a new control around the selection or as a placeholder', () => {
		const state = stateOf([{ text: 'Plain words' }]);
		const selected = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 6)));
		const wrapped = selected.apply(insertContentControl(selected, 'richText', { alias: 'Wrap' })!);
		expect(controlSpans(wrapped.doc)[0]).toMatchObject({ from: 1, to: 6 });
		const picker = state.apply(insertContentControl(state, 'date')!);
		expect(controlSpans(picker.doc)[0]?.control).toMatchObject({
			kind: 'date',
			showingPlaceholder: true,
			dateFormat: 'M/d/yyyy',
		});
		expect(text(picker)).toContain('Click or tap to enter a date.');
	});
});

describe('content control locks', () => {
	const locked = (lock: ContentControl['lock']) =>
		stateOf([
			{ text: 'A ' },
			{ text: 'locked', contentControl: control({ kind: 'text', ...(lock ? { lock } : {}) }) },
		]);

	it('contentLocked refuses edits inside but allows removing the whole control', () => {
		const state = locked('contentLocked');
		expect(violatesControlLock(state.tr.insertText('x', 5), state)).toBe(true);
		expect(violatesControlLock(state.tr.delete(4, 6), state)).toBe(true);
		expect(violatesControlLock(state.tr.insertText('x', 1), state)).toBe(false);
		expect(violatesControlLock(state.tr.delete(3, 9), state)).toBe(false);
		expect(state.applyTransaction(state.tr.insertText('x', 5)).transactions).toHaveLength(0);
	});

	it('sdtLocked refuses removing the control but allows editing inside', () => {
		const state = locked('sdtLocked');
		expect(violatesControlLock(state.tr.delete(3, 9), state)).toBe(true);
		expect(violatesControlLock(state.tr.insertText('x', 5), state)).toBe(false);
		expect(violatesControlLock(state.tr.delete(4, 6), state)).toBe(false);
	});

	it('a content-locked block refuses typing anywhere in it, its ends included', () => {
		const blockControl = control({ level: 'block', lock: 'contentLocked' });
		const doc = schema.node('doc', null, [
			schema.node('paragraph', { contentControl: blockControl }, [schema.text('Notes')]),
			schema.node('paragraph', null, [schema.text('After')]),
		]);
		const state = EditorState.create({ schema, doc });
		for (const at of [1, 3, 6])
			expect(violatesControlLock(state.tr.insertText('x', at), state)).toBe(true);
		expect(violatesControlLock(state.tr.insertText('x', 9), state)).toBe(false);
		// Deleting the whole block is not editing its content.
		expect(violatesControlLock(state.tr.delete(0, 7), state)).toBe(false);
	});

	it('finds the control around a position', () => {
		const state = locked(undefined);
		const spans = controlSpans(state.doc);
		expect(controlAt(spans, 5)?.control.id).toBe('11');
		expect(controlAt(spans, 1)).toBeUndefined();
	});
});
