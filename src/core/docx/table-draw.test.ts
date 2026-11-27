import { describe, expect, it } from 'vitest';
import type { Table, TableCell } from './model';
import { tableGrid } from './table-grid';
import { drawTableLine, drawnTableBox, eraseTableEdge } from './table-draw';
import type { IdFactory } from './table-structure-cells';
import { twips } from './units';

let serial = 0;
const newId: IdFactory = () => `new-${++serial}`;

const cell = (text: string, extra: Partial<TableCell> = {}): TableCell => ({
	paragraphs: [{ type: 'paragraph', id: `p-${text || ++serial}`, runs: [{ text }] }],
	...extra,
});

const plain = (rows: number, columns: number, width = 1440): Table => ({
	type: 'table',
	id: 't',
	grid: Array.from({ length: columns }, () => twips(width)),
	rows: Array.from({ length: rows }, (_, r) =>
		Array.from({ length: columns }, (_, c) => cell(`r${r}c${c}`)),
	),
});

/** Each row as `text:span` strings, `^` for a merge start and `|` for a continuation. */
const shape = (table: Table) =>
	table.rows.map((row) =>
		row.map(
			(c) =>
				`${c.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('/')}:${c.gridSpan ?? 1}${
					c.verticalMerge === 'restart' ? '^' : c.verticalMerge === 'continue' ? '|' : ''
				}`,
		),
	);

describe('drawTableLine', () => {
	it('splits a cell with a vertical line, cutting the grid column under it', () => {
		const result = drawTableLine(
			plain(2, 2),
			{ orientation: 'vertical', row: 0, index: 0, fraction: 0.25 },
			newId,
		);
		expect(result?.grid).toEqual([360, 1080, 1440]);
		expect(shape(result!)).toEqual([
			['r0c0:1', ':1', 'r0c1:1'],
			['r1c0:2', 'r1c1:1'],
		]);
		expect(tableGrid(result!)).toBeDefined();
	});

	it('snaps a vertical line to a column boundary inside a merged cell', () => {
		const table = plain(2, 2);
		table.rows[0] = [cell('wide', { gridSpan: 2 })];
		const result = drawTableLine(
			table,
			{ orientation: 'vertical', row: 0, index: 0, fraction: 0.51 },
			newId,
		);
		expect(result?.grid).toEqual([1440, 1440]);
		expect(shape(result!)[0]).toEqual(['wide:1', ':1']);
	});

	it('refuses a line drawn at the very edge of a cell', () => {
		expect(
			drawTableLine(
				plain(1, 1),
				{ orientation: 'vertical', row: 0, index: 0, fraction: 0.01 },
				newId,
			),
		).toBeUndefined();
	});

	it('builds an even grid from the default width when the table has none', () => {
		const table = plain(1, 2);
		delete table.grid;
		const result = drawTableLine(
			table,
			{ orientation: 'vertical', row: 0, index: 1, fraction: 0.5 },
			newId,
			{ defaultWidthTwips: 2880 },
		);
		expect(result?.grid).toEqual([1440, 720, 720]);
		expect(result?.rows[0]).toHaveLength(3);
	});

	it('splits every row of a vertical merge with a vertical line', () => {
		const table = plain(2, 1);
		table.rows[0]![0]!.verticalMerge = 'restart';
		table.rows[1]![0]!.verticalMerge = 'continue';
		const result = drawTableLine(
			table,
			{ orientation: 'vertical', row: 0, index: 0, fraction: 0.5 },
			newId,
		);
		expect(shape(result!).map((row) => row.map((c) => c.split(':')[1]))).toEqual([
			['1^', '1^'],
			['1|', '1|'],
		]);
	});

	it('splits a cell into two rows with a horizontal line', () => {
		const result = drawTableLine(
			plain(1, 2),
			{ orientation: 'horizontal', row: 0, index: 1, fraction: 0.5 },
			newId,
		);
		expect(result?.rows).toHaveLength(2);
		expect(shape(result!)[0]).toEqual(['r0c0:1^', 'r0c1:1']);
		expect(shape(result!)[1]?.[0]).toBe(':1|');
	});

	it('takes a vertical merge apart at the nearest row boundary', () => {
		const table = plain(3, 2);
		table.rows[0]![0]!.verticalMerge = 'restart';
		table.rows[1]![0]!.verticalMerge = 'continue';
		table.rows[2]![0]!.verticalMerge = 'continue';
		const result = drawTableLine(
			table,
			{ orientation: 'horizontal', row: 0, index: 0, fraction: 0.7 },
			newId,
		);
		expect(result?.rows.map((row) => row[0]?.verticalMerge)).toEqual([
			'restart',
			'continue',
			undefined,
		]);
	});
});

describe('eraseTableEdge', () => {
	it('merges the cells on either side of an inner vertical line', () => {
		const outcome = eraseTableEdge(plain(2, 2), { row: 0, index: 0, side: 'right' }, newId);
		expect(outcome.kind).toBe('merged');
		if (outcome.kind !== 'merged') return;
		expect(shape(outcome.table)).toEqual([['r0c0/r0c1:2'], ['r1c0:1', 'r1c1:1']]);
	});

	it('merges the cells above and below an inner horizontal line', () => {
		const outcome = eraseTableEdge(plain(2, 2), { row: 1, index: 1, side: 'top' }, newId);
		if (outcome.kind !== 'merged') throw new Error(outcome.kind);
		expect(outcome.table.rows[0]?.[1]?.verticalMerge).toBe('restart');
		expect(outcome.table.rows[1]?.[1]?.verticalMerge).toBe('continue');
	});

	it('reports an outer line', () => {
		expect(eraseTableEdge(plain(2, 2), { row: 0, index: 0, side: 'left' }, newId).kind).toBe(
			'outer',
		);
		expect(eraseTableEdge(plain(2, 2), { row: 1, index: 1, side: 'bottom' }, newId).kind).toBe(
			'outer',
		);
	});

	it('refuses when the two cells do not form a rectangle', () => {
		const table = plain(2, 2);
		table.rows[0]![0]!.verticalMerge = 'restart';
		table.rows[1]![0]!.verticalMerge = 'continue';
		expect(eraseTableEdge(table, { row: 0, index: 0, side: 'right' }, newId).kind).toBe('refused');
	});

	it('picks the neighbour under the pointer along a merged side', () => {
		const table = plain(2, 2);
		table.rows[0]![1]!.verticalMerge = 'restart';
		table.rows[1]![1]!.verticalMerge = 'continue';
		const outcome = eraseTableEdge(
			table,
			{ row: 0, index: 0, side: 'bottom', fraction: 0.5 },
			newId,
		);
		expect(outcome.kind).toBe('merged');
	});
});

describe('drawnTableBox', () => {
	it('clamps to the column width and refuses a tiny rectangle', () => {
		expect(drawnTableBox(20000, 600, 9360)).toEqual({ widthTwips: 9360, heightTwips: 600 });
		expect(drawnTableBox(100, 600)).toBeUndefined();
	});
});
