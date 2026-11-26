import { expect, test } from 'vitest';
import { runToInlineNodes } from 'ooxml-core/docx/ui';
import type { TextRun } from 'ooxml-core/docx';
import { schema } from '../schema';
import { markAllPositions } from './mark-all';

test('marks the first exact whole word and skips simple and multi-paragraph field results', () => {
	const paragraph = (...runs: TextRun[]) =>
		schema.nodes.paragraph!.create(
			null,
			runs.flatMap((run) => runToInlineNodes(run, schema)),
		);
	const literal = paragraph({ text: 'zebra Zebras Zebra ' }, { text: 'Zebra Zebra', bold: true });
	const start = paragraph({ text: '', fieldChar: 'begin' }, { text: '', fieldChar: 'separate' });
	const result = paragraph({ text: 'Zebra Zebra' });
	const end = paragraph(
		{ text: '', fieldChar: 'end' },
		{ text: 'Zebra', field: { instr: 'QUOTE Zebra', simple: true } },
		{ text: ' Zebra Zebra' },
	);
	const doc = schema.nodes.doc!.create(null, [literal, start, result, end]);
	expect(markAllPositions(doc, 'Zebra')).toEqual([
		1 + 'zebra Zebras Zebra'.length,
		1 + literal.nodeSize + start.nodeSize + result.nodeSize + 1 + end.child(1).nodeSize + 6,
	]);
});
