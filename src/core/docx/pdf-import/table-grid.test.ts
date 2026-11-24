import { describe, expect, it } from 'vitest';
import type { PdfPageRule } from '../../pdf/index';
import { findGrids } from './table-grid';

const line = (x0: number, y0: number, x1: number, y1: number): PdfPageRule => ({
	x0,
	y0,
	x1,
	y1,
	kind: 'line',
	stroked: true,
	filled: false,
	color: '#000000',
	lineWidth: 0.75,
});
const fill = (x0: number, y0: number, x1: number, y1: number, color: string): PdfPageRule => ({
	x0,
	y0,
	x1,
	y1,
	kind: 'rect',
	stroked: false,
	filled: true,
	color,
	lineWidth: 1,
});

describe('findGrids', () => {
	it('builds the grid, merges where a rule is missing and finds the shading', () => {
		// Three columns (100, 200, 300, 400) and two rows (100, 120, 140); the second row has no
		// rule between its last two columns.
		const rules = [
			fill(100, 100, 200, 120, '#ffcc00'),
			line(100, 100, 400, 100),
			line(100, 120, 400, 120),
			line(100, 140, 400, 140),
			line(100, 100, 100, 140),
			line(200, 100, 200, 140),
			line(300, 100, 300, 120),
			line(400, 100, 400, 140),
			line(120, 300, 180, 300),
		];
		const { grids, rest } = findGrids(rules);
		expect(grids).toHaveLength(1);
		const grid = grids[0]!;
		expect(grid.xs.map(Math.round)).toEqual([100, 200, 300, 400]);
		expect(grid.ys.map(Math.round)).toEqual([100, 120, 140]);
		expect(grid.cells.map((cell) => [cell.row, cell.col, cell.rowSpan, cell.colSpan])).toEqual([
			[0, 0, 1, 1],
			[0, 1, 1, 1],
			[0, 2, 1, 1],
			[1, 0, 1, 1],
			[1, 1, 1, 2],
		]);
		expect(grid.cells[0]!.fill).toBe('#ffcc00');
		expect(grid.cells[0]!.borders.top).toEqual({ width: 0.75, color: '#000000' });
		// The rule far below the table is an underline, not part of it.
		expect(rest).toEqual([rules[8]]);
	});

	it('merges vertically, joins rules drawn cell by cell and ignores a lone box', () => {
		const pieces = [
			...[100, 120, 140].flatMap((y) => [line(100, y, 200, y), line(200, y, 300, y)]),
			line(100, 100, 100, 140),
			line(300, 100, 300, 140),
			line(200, 100, 200, 140),
		];
		// Remove the rule under the second column's first row: its two units become one cell.
		pieces.splice(3, 1);
		const { grids } = findGrids(pieces);
		expect(grids[0]!.cells.map((cell) => [cell.row, cell.col, cell.rowSpan, cell.colSpan])).toEqual(
			[
				[0, 0, 1, 1],
				[0, 1, 2, 1],
				[1, 0, 1, 1],
			],
		);
		const box = [
			line(50, 50, 500, 50),
			line(50, 700, 500, 700),
			line(50, 50, 50, 700),
			line(500, 50, 500, 700),
		];
		expect(findGrids(box).grids).toEqual([]);
	});
});
