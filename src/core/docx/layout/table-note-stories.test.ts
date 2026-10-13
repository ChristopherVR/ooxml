import { expect, it } from 'vitest';
import { independentNoteStories } from './table-note-stories';
import type { LayoutTable, LayoutTableCell } from './input';
const create = (): LayoutTableCell => ({
	paragraphs: [],
	blocks: [
		{
			kind: 'table',
			id: 'inner',
			rows: [
				{
					cells: [
						{ paragraphs: [{ kind: 'paragraph', id: 'a', runs: [] }] },
						{ paragraphs: [{ kind: 'paragraph', id: 'b', runs: [] }] },
					],
				},
			],
		},
		{ kind: 'paragraph', id: 'tail', runs: [] },
	],
});
it('groups independent leaf stories without including the sequential trailing paragraph', () => {
	const groups = independentNoteStories(create());
	expect([...groups!]).toEqual([
		['a', 'inner:0'],
		['b', 'inner:1'],
	]);
});
it.each(['center', 'exact', 'multiple', 'deep', 'header', 'leading', 'keep'] as const)(
	'refuses unmeasured independent nested profile: %s',
	(kind) => {
		const cell = create(),
			table = cell.blocks![0] as LayoutTable;
		if (kind === 'center') table.rows[0]!.cells[0]!.verticalAlign = 'center';
		if (kind === 'exact') {
			table.rows[0]!.heightPx = 100;
			table.rows[0]!.heightRule = 'exact';
		}
		if (kind === 'multiple') table.rows.push({ ...table.rows[0]! });
		if (kind === 'deep') table.rows[0]!.cells[0]!.blocks = [table];
		if (kind === 'keep') table.rows[0]!.cells[0]!.paragraphs[0]!.keepLines = true;
		if (kind === 'header') table.rows[0]!.isHeader = true;
		if (kind === 'leading') cell.blocks!.unshift({ kind: 'paragraph', id: 'lead', runs: [] });
		expect(independentNoteStories(cell)).toBeUndefined();
	},
);
