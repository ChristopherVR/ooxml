import { expect, it } from 'vitest';
import { Fragment, Schema, Slice } from 'prosemirror-model';
import { fieldClipboardSlice } from './field-clipboard';
import { markSpecs } from './schema-marks';
import {
	fieldMarkerNodeSpec,
	hardBreakNodeSpec,
	pageBreakNodeSpec,
	noteReferenceNodeSpec,
} from './break-note-schema';
import { runToInlineNodes, inlineNodeRun } from './run-adapter';
const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*' },
		text: { group: 'inline' },
		fieldMarker: fieldMarkerNodeSpec,
		hardBreak: hardBreakNodeSpec,
		pageBreak: pageBreakNodeSpec,
		noteReference: noteReferenceNodeSpec,
	},
	marks: markSpecs,
});
const result = runToInlineNodes(
	{ text: 'AB', field: { instr: 'REF Bookmark' }, bold: true },
	schema,
)[0]!;
const marker = (kind: 'begin' | 'end') => schema.node('fieldMarker', { kind });
it.each([
	{ text: '\n' },
	{ text: '', break: 'page' as const },
	{ text: '', noteReference: { kind: 'footnote' as const, id: '7' } },
])(
	'projects standalone cached inline objects without losing their own properties: %j',
	(content) => {
		const atom = runToInlineNodes(
			{
				...content,
				bold: true,
				field: { instr: 'REF Target', simple: true },
				fieldInstanceId: 'source',
			},
			schema,
		)[0]!;
		const copied = fieldClipboardSlice(new Slice(Fragment.from(atom), 0, 0)).content.firstChild!;
		expect(inlineNodeRun(copied)).toEqual({ ...content, bold: true });
		expect(inlineNodeRun(atom)?.field).toEqual({ instr: 'REF Target', simple: true });
	},
);
it('retains a complete complex field cached line break and its effective formatting', () => {
	const atom = runToInlineNodes(
		{ text: '\n', bold: true, field: { instr: 'REF Target' } },
		schema,
	)[0]!;
	const effective = atom.type.create(
		{ ...atom.attrs, runFormat_bold: 'false' },
		undefined,
		atom.marks,
	);
	const slice = new Slice(Fragment.fromArray([marker('begin'), effective, marker('end')]), 0, 0);
	const copied = fieldClipboardSlice(slice);
	expect(inlineNodeRun(copied.content.child(1))).toMatchObject({
		text: '\n',
		bold: false,
		field: { instr: 'REF Target' },
	});
	const literal = fieldClipboardSlice(new Slice(Fragment.from(effective), 0, 0));
	expect(inlineNodeRun(literal.content.firstChild!)).toEqual({ text: '\n', bold: false });
});
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
