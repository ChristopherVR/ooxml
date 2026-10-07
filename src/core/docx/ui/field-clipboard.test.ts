import { expect, it } from 'vitest';
import { Fragment, Schema, Slice } from 'prosemirror-model';
import { fieldClipboardSlice } from './field-clipboard';
import { markSpecs } from './schema-marks';
import { fieldMarkerNodeSpec } from './break-note-schema';
import { runToInlineNodes } from './run-adapter';
const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*' },
		text: { group: 'inline' },
		fieldMarker: fieldMarkerNodeSpec,
	},
	marks: markSpecs,
});
const result = runToInlineNodes(
	{ text: 'AB', field: { instr: 'REF Bookmark' }, bold: true },
	schema,
)[0]!;
const marker = (kind: 'begin' | 'end') => schema.node('fieldMarker', { kind });
it('strips field metadata from an incomplete complex field without changing slice openness', () => {
	const doc = schema.node('doc', null, schema.node('paragraph', null, [marker('begin'), result]));
	const slice = new Slice(doc.content, 1, 1);
	const copied = fieldClipboardSlice(slice);
	expect(copied.openStart).toBe(1);
	expect(copied.openEnd).toBe(1);
	expect(copied.content.firstChild!.childCount).toBe(1);
	const marks = copied.content.firstChild!.lastChild!.marks.map((mark) => mark.type.name);
	expect(marks).toEqual(['bold']);
	expect(doc.firstChild!.lastChild!.marks.some((mark) => mark.type.name === 'field')).toBe(true);
});

it('drops unmatched result separators, instruction runs and end markers from copied text', () => {
	const paragraph = schema.node('paragraph', null, [
		schema.node('fieldMarker', { kind: 'code', code: 'REF Secret' }),
		schema.node('fieldMarker', { kind: 'separate' }),
		result,
		marker('end'),
		schema.text('R'),
	]);
	const copied = fieldClipboardSlice(new Slice(Fragment.from(paragraph), 1, 1));
	expect(copied.content.firstChild!.textContent).toBe('ABR');
	expect(copied.content.firstChild!.childCount).toBe(2);
});

it('retains a complete inner field inside an incomplete outer field', () => {
	const doc = schema.node(
		'doc',
		null,
		schema.node('paragraph', null, [
			marker('begin'),
			marker('begin'),
			result,
			marker('end'),
			result,
		]),
	);
	const copied = fieldClipboardSlice(new Slice(doc.content, 1, 1));
	const paragraph = copied.content.firstChild!;
	expect(paragraph.childCount).toBe(4);
	expect(paragraph.child(0).attrs.kind).toBe('begin');
	expect(paragraph.child(1).marks.some((mark) => mark.type.name === 'field')).toBe(true);
	expect(paragraph.child(2).attrs.kind).toBe('end');
	expect(paragraph.child(3).marks.map((mark) => mark.type.name)).toEqual(['bold']);
});
it('removes structural markers from a complete-looking field with invalid marker ordering', () => {
	const doc = schema.node(
		'doc',
		null,
		schema.node('paragraph', null, [
			marker('begin'),
			schema.node('fieldMarker', { kind: 'separate' }),
			schema.node('fieldMarker', { kind: 'code', code: 'REF Target' }),
			result,
			marker('end'),
		]),
	);
	const copied = fieldClipboardSlice(new Slice(doc.content, 1, 1));
	expect(copied.content.firstChild!.childCount).toBe(1);
	expect(copied.content.firstChild!.firstChild!.marks.map((mark) => mark.type.name)).toEqual([
		'bold',
	]);
});
it('retains complete nested complex fields without retaining an unrelated result-only run', () => {
	const doc = schema.node(
		'doc',
		null,
		schema.node('paragraph', null, [
			marker('begin'),
			marker('begin'),
			result,
			marker('end'),
			marker('end'),
			result,
		]),
	);
	const copied = fieldClipboardSlice(new Slice(doc.content, 0, 0));
	expect(copied.content.firstChild!.child(2).marks.some((mark) => mark.type.name === 'field')).toBe(
		true,
	);
	expect(copied.content.firstChild!.lastChild!.marks.map((mark) => mark.type.name)).toEqual([
		'bold',
	]);
});
