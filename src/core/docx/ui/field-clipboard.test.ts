import { expect, it } from 'vitest';
import { Schema, Slice } from 'prosemirror-model';
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
	const marks = copied.content.firstChild!.lastChild!.marks.map((mark) => mark.type.name);
	expect(marks).toEqual(['bold']);
	expect(doc.firstChild!.lastChild!.marks.some((mark) => mark.type.name === 'field')).toBe(true);
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
