import { describe, expect, it } from 'vitest';
import { createDocument } from './index';
import type { Paragraph } from './model';
import { numberMoveNames } from './move-names';
import { sectionsOf } from './section-layout';

const paragraph = (
	id: string,
	name: string,
	kind: 'moveFrom' | 'moveTo' = 'moveFrom',
): Paragraph => ({
	type: 'paragraph',
	id,
	runs: [{ text: id, revision: { id, kind, author: 'Ada', move: { name } } }],
});
const nameOf = (block: Paragraph) => block.runs[0]!.revision!.move!.name;

describe('Word move name export', () => {
	it('reserves existing numeric names and renumbers both sides without changing internal identity', () => {
		const model = createDocument();
		model.blocks = [
			paragraph('a', 'move9'),
			paragraph('b', 'move-client-a'),
			paragraph('c', 'move-client-a', 'moveTo'),
			paragraph('d', 'move-client-b'),
		];
		const before = structuredClone(model);
		const numbered = numberMoveNames(model);
		expect(numbered.blocks.map((block) => nameOf(block as Paragraph))).toEqual([
			'move9',
			'move10',
			'move10',
			'move11',
		]);
		expect(model).toEqual(before);
		expect(numberMoveNames(numbered)).toBe(numbered);
	});
	it('shares one numbering map across body, table cells, section stories and notes', () => {
		const model = createDocument();
		model.blocks = [
			paragraph('a', 'internal'),
			{
				type: 'table',
				id: 'table',
				rows: [[{ paragraphs: [paragraph('cell', 'internal', 'moveTo')] }]],
			},
		];
		model.sections = [
			{
				...sectionsOf(model)[0]!,
				headers: { default: { blocks: [paragraph('header', 'other')] } },
				footers: { default: { blocks: [paragraph('footer', 'other', 'moveTo')] } },
			},
		];
		model.footnotes = [{ id: '1', blocks: [paragraph('note', 'internal')] }];
		model.endnotes = [{ id: '2', blocks: [paragraph('endnote', 'other')] }];
		const numbered = numberMoveNames(model);
		expect(nameOf(numbered.blocks[0] as Paragraph)).toBe('move1');
		const table = numbered.blocks[1]!;
		if (table.type !== 'table') throw new Error('Expected table');
		expect(nameOf(table.rows[0]![0]!.paragraphs[0]!)).toBe('move1');
		expect(nameOf(numbered.sections![0]!.headers!.default!.blocks[0] as Paragraph)).toBe('move2');
		expect(nameOf(numbered.sections![0]!.footers!.default!.blocks[0] as Paragraph)).toBe('move2');
		expect(nameOf(numbered.footnotes![0]!.blocks[0] as Paragraph)).toBe('move1');
		expect(nameOf(numbered.endnotes![0]!.blocks[0] as Paragraph)).toBe('move2');
	});
});
