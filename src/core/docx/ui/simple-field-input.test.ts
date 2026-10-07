import { expect, it } from 'vitest';
import { Fragment, Schema, Slice } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import { markSpecs } from './schema-marks';
import { runToInlineNodes, inlineNodeRun } from './run-adapter';
import { fieldResultRanges } from './field-results';
import { fieldClipboardSlice } from './field-clipboard';
import {
	deleteSimpleFieldResult,
	replaceSimpleFieldResult,
	simpleFieldPasteSlice,
} from './simple-field-input';
import { fieldMarkerNodeSpec } from './break-note-schema';
import { deleteFieldSelection, fieldsBalanced } from './field-guard';
const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*' },
		text: { group: 'inline' },
		fieldMarker: fieldMarkerNodeSpec,
	},
	marks: markSpecs,
});
it('retains an empty instruction independently from an adjacent identical simple field', () => {
	const state = EditorState.create({ doc, selection: TextSelection.create(doc, 2, 4) });
	const tr = deleteSimpleFieldResult(state, true)!;
	expect(tr.doc.textContent).toBe('LCDR');
	expect(tr.selection.from).toBe(5);
	expect(fieldResultRanges(tr.doc).map((run) => run.text)).toEqual(['CD']);
	const runs = Array.from({ length: tr.doc.firstChild!.childCount }, (_, index) =>
		inlineNodeRun(tr.doc.firstChild!.child(index))!,
	);
	expect(runs.filter((run) => run.fieldChar).map((run) => run.fieldChar)).toEqual([
		'begin',
		'separate',
		'end',
	]);
	expect(runs.find((run) => run.fieldCode)?.fieldCode).toBe('REF Target');
	expect(runs.find((run) => run.field)?.fieldInstanceId).toBe('second');
	expect(replaceSimpleFieldResult(state, 2, 3, '')).toBeNull();
});
const field = (text: string, id: string) =>
	runToInlineNodes(
		{ text, field: { instr: 'REF Target', simple: true }, fieldInstanceId: id },
		schema,
	);
const doc = schema.node(
	'doc',
	null,
	schema.node('paragraph', null, [
		schema.text('L'),
		...field('AB', 'first'),
		...field('CD', 'second'),
		schema.text('R'),
	]),
);
it('shares field-aware deletion while leaving ordinary and whole-field selections to the host', () => {
	const complex = schema.node(
		'doc',
		null,
		schema.node('paragraph', null, [
			schema.text('L'),
			...[
				{ text: '', fieldChar: 'begin' as const },
				{ text: '', fieldCode: 'REF Target' },
				{ text: '', fieldChar: 'separate' as const },
				{ text: 'AB', field: { instr: 'REF Target' } },
				{ text: '', fieldChar: 'end' as const },
				{ text: 'R' },
			].flatMap((run) => runToInlineNodes(run, schema)),
		]),
	);
	const state = (from: number, to: number) =>
		EditorState.create({ doc: complex, selection: TextSelection.create(complex, from, to) });
	const tr = deleteFieldSelection(state(6, 8))!;
	expect(tr.doc.textContent).toBe('LAR');
	expect(fieldsBalanced(tr.doc)).toBe(true);
	expect(deleteFieldSelection(state(2, 8))).toBeNull();
	expect(deleteFieldSelection(state(1, 2))).toBeNull();
	expect(deleteFieldSelection(state(5, 5))).toBeNull();
});
it('retains full-result replacements without joining adjacent identical fields', () => {
	const state = EditorState.create({ doc });
	const tr = replaceSimpleFieldResult(state, 2, 4, 'X')!;
	expect(tr.doc.textContent).toBe('LXCDR');
	expect(fieldResultRanges(tr.doc).map((run) => run.text)).toEqual(['X', 'CD']);
	expect(inlineNodeRun(tr.doc.nodeAt(2)!)!.fieldInstanceId).toBe('first');
	expect(replaceSimpleFieldResult(state, 4, 4, 'X')).toBeNull();
	expect(replaceSimpleFieldResult(state, 2, 2, 'X')).toBeNull();
});
it('adopts inline paste into the target field while retaining source formatting', () => {
	let state = EditorState.create({ doc, selection: TextSelection.create(doc, 3) });
	const source = runToInlineNodes(
		{
			text: 'X',
			field: { instr: 'REF Foreign', simple: true },
			fieldInstanceId: 'foreign',
			smallCaps: true,
		},
		schema,
	);
	const pasted = simpleFieldPasteSlice(
		fieldClipboardSlice(new Slice(Fragment.fromArray(source), 0, 0)),
		state,
	);
	expect(inlineNodeRun(pasted.content.firstChild!)!).toMatchObject({
		text: 'X',
		field: { instr: 'REF Target', simple: true },
		fieldInstanceId: 'first',
		smallCaps: true,
	});
	state = state.apply(state.tr.replaceSelection(pasted));
	expect(fieldResultRanges(state.doc).map((run) => run.text)).toEqual(['AXB', 'CD']);
});
it('leaves closed structural slices and boundary cursors outside result adoption', () => {
	const state = EditorState.create({ doc, selection: TextSelection.create(doc, 3) });
	const slice = new Slice(Fragment.from(schema.node('paragraph', null, schema.text('X'))), 0, 0);
	expect(simpleFieldPasteSlice(slice, state)).toBe(slice);
	const inline = new Slice(Fragment.from(schema.text('X')), 0, 0);
	expect(
		simpleFieldPasteSlice(
			inline,
			EditorState.create({ doc, selection: TextSelection.create(doc, 4) }),
		),
	).toBe(inline);
});
