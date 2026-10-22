import { describe, expect, it } from 'vitest';
import { formatRange } from '../address';
import { changeAreas } from './change-areas';
import type { CellPosition } from './engine-types';

const block = (sheet: number, rows: number, cols: number, row0 = 0, col0 = 0): CellPosition[] =>
	Array.from({ length: rows * cols }, (_, i) => ({
		sheet,
		row: row0 + Math.floor(i / cols),
		col: col0 + (i % cols),
	}));

/** Every cell the areas cover, as `sheet:A1` keys. */
const covered = (changes: CellPosition[]) =>
	new Set(
		changeAreas(changes).flatMap(({ sheet, range }) => {
			const keys: string[] = [];
			for (let row = range.start.row; row <= range.end.row; row++)
				for (let col = range.start.col; col <= range.end.col; col++)
					keys.push(`${sheet}:${row},${col}`);
			return keys;
		}),
	);

describe('changeAreas', () => {
	it('turns a row-major block into one rectangle', () => {
		const areas = changeAreas(block(0, 1000, 10, 5, 2));
		expect(areas.map((a) => [a.sheet, formatRange(a.range)])).toEqual([[0, 'C6:L1005']]);
	});

	it('covers exactly the changed cells whatever their order, sheets or duplicates', () => {
		const changes = [
			...block(0, 3, 4),
			...block(1, 2, 2, 10, 10),
			{ sheet: 0, row: 1, col: 1 },
			{ sheet: 0, row: 7, col: 0 },
			...block(0, 2, 4).reverse(),
		];
		const expected = new Set(changes.map((c) => `${c.sheet}:${c.row},${c.col}`));
		expect(covered(changes)).toEqual(expected);
		expect(changeAreas([])).toEqual([]);
	});
});
