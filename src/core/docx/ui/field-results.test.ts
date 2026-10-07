import { expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { fieldResultRanges } from './field-results';
import { commentSelectionRange } from './comment-selection';
import { markSpecs } from './schema-marks';
import { runToInlineNodes } from './run-adapter';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*' },
		text: { group: 'inline' },
	},
	marks: markSpecs,
});
const field = schema.marks.field!.create({ instr: 'QUOTE ABCDE', simple: true });
const doc = schema.node('doc', null, [
	schema.node('paragraph', null, [
		schema.text('L'),
		schema.text('AB', [field]),
		schema.text('CDE', [field, schema.marks.bold!.create()]),
		schema.text('R'),
	]),
	schema.node('paragraph', null, schema.text('ABCDE', [field])),
]);
it('groups split result formatting while retaining paragraph boundaries and replacement marks', () => {
	const ranges = fieldResultRanges(doc);
	expect(ranges.map(({ from, to, text }) => ({ from, to, text }))).toEqual([
		{ from: 2, to: 7, text: 'ABCDE' },
		{ from: 10, to: 15, text: 'ABCDE' },
	]);
	expect(ranges[0]!.marks).toEqual([field]);
});
it('expands simple field comments across result formatting, without touching adjacent text', () => {
	expect(commentSelectionRange(doc, 4, 5)).toEqual({ from: 2, to: 7 });
	expect(commentSelectionRange(doc, 1, 2)).toEqual({ from: 1, to: 2 });
	expect(commentSelectionRange(doc, 4, 4)).toEqual({ from: 4, to: 4 });
});

it.each([
	[undefined, { locked: false }],
	[
		{ locked: false, dirty: true },
		{ locked: true, dirty: true },
	],
	[
		{ locked: true, dirty: false },
		{ locked: true, dirty: true },
	],
] as const)(
	'keeps anonymous adjacent field caches with unequal flags separate (%j, %j)',
	(first, second) => {
		const runs = [first, second].flatMap((fieldFlags, index) =>
			runToInlineNodes(
				{
					text: String(index),
					field: { instr: 'REF Target', simple: true },
					...(fieldFlags && { fieldFlags }),
				},
				schema,
			),
		);
		const adjacent = schema.node('doc', null, schema.node('paragraph', null, runs));
		expect(fieldResultRanges(adjacent).map((range) => range.text)).toEqual(['0', '1']);
	},
);
