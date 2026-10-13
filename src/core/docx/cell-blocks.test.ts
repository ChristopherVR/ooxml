import { expect, it } from 'vitest';
import type { Paragraph, Table, TableCell } from './model';
import { cellSourceBlocks } from './cell-blocks';
const paragraph = (id: string): Paragraph => ({ type: 'paragraph', id, runs: [{ text: id }] });
const table = (id: string): Table => ({
	type: 'table',
	id,
	rows: [[{ paragraphs: [paragraph(id + 'p')] }]],
});
it('keeps nested source tables anchored after new paragraphs beside each original direct paragraph', () => {
	const source: TableCell = {
		paragraphs: ['a', 'b', 'c'].map(paragraph),
		sourceBlocks: [paragraph('a'), table('n0'), paragraph('b'), table('n1'), paragraph('c')],
	};
	source.paragraphs = ['a', 'new-a', 'b', 'new-b', 'c'].map(paragraph);
	expect(cellSourceBlocks(source).map((block) => block.id)).toEqual([
		'a',
		'new-a',
		'n0',
		'b',
		'new-b',
		'n1',
		'c',
	]);
});
it('keeps leading and trailing inserted paragraphs while retaining a nested table before its direct paragraph', () => {
	const source: TableCell = {
		paragraphs: ['leading', 'a', 'trailing'].map(paragraph),
		sourceBlocks: [table('n0'), paragraph('a')],
	};
	expect(cellSourceBlocks(source).map((block) => block.id)).toEqual([
		'n0',
		'leading',
		'a',
		'trailing',
	]);
});
