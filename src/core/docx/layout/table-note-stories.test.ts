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

const blocksOnly = (): LayoutTableCell => {
	const cell = create(),
		table = cell.blocks![0] as LayoutTable;
	for (const inner of table.rows[0]!.cells) {
		inner.blocks = inner.paragraphs;
		inner.paragraphs = [];
	}
	return cell;
};
it('groups the actual blocks-only paragraph stories used by source layout', () => {
	expect([...independentNoteStories(blocksOnly())!]).toEqual([
		['a', 'inner:0'],
		['b', 'inner:1'],
	]);
});
it.each(['keepLines', 'keepNext', 'pageBreakBefore'] as const)(
	'refuses blocks-only story with %s',
	(property) => {
		const cell = blocksOnly(),
			table = cell.blocks![0] as LayoutTable;
		const paragraph = table.rows[0]!.cells[0]!.blocks![0]!;
		if (paragraph.kind !== 'paragraph') throw new Error('Expected paragraph');
		paragraph[property] = true;
		expect(independentNoteStories(cell)).toBeUndefined();
	},
);
it('refuses an empty story map when authoritative blocks omit cached paragraphs', () => {
	const cell = create(),
		table = cell.blocks![0] as LayoutTable;
	for (const inner of table.rows[0]!.cells) inner.blocks = [];
	expect(independentNoteStories(cell)).toBeUndefined();
});
it('does not admit cached paragraph IDs absent from authoritative source blocks', () => {
	const cell = blocksOnly(),
		table = cell.blocks![0] as LayoutTable;
	table.rows[0]!.cells[0]!.paragraphs = [
		{ kind: 'paragraph', id: 'stale', runs: [], keepLines: true },
	];
	expect([...independentNoteStories(cell)!]).toEqual([
		['a', 'inner:0'],
		['b', 'inner:1'],
	]);
});

it('admits one actual marker without double counting its associated note metadata', () => {
	const cell = blocksOnly(),
		table = cell.blocks![0] as LayoutTable;
	const paragraph = table.rows[0]!.cells[0]!.blocks![0]!;
	if (paragraph.kind !== 'paragraph') throw new Error('Expected paragraph');
	paragraph.runs = [{ text: '1', footnoteReferenceId: 'note1' }];
	paragraph.footnotes = [{ id: 'note1', paragraphs: [] }];
	expect(independentNoteStories(cell)?.size).toBe(2);
});
it.each(['distinct', 'reused', 'unanchored'] as const)(
	'refuses multiple interacting inner references: %s',
	(kind) => {
		const cell = blocksOnly(),
			table = cell.blocks![0] as LayoutTable;
		const a = table.rows[0]!.cells[0]!.blocks![0]!,
			b = table.rows[0]!.cells[1]!.blocks![0]!;
		if (a.kind !== 'paragraph' || b.kind !== 'paragraph') throw new Error('Expected paragraph');
		a.runs = [{ text: '1', footnoteReferenceId: 'note1' }];
		a.footnotes = [{ id: 'note1', paragraphs: [] }];
		if (kind === 'unanchored') b.footnotes = [{ id: 'note2', paragraphs: [] }];
		else b.runs = [{ text: '2', footnoteReferenceId: kind === 'reused' ? 'note1' : 'note2' }];
		expect(independentNoteStories(cell)).toBeUndefined();
	},
);
