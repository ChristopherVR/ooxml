// The cells a step lists for recalculation: each once and in the same order as before.
import { describe, expect, it } from 'vitest';
import { cellKey, type CellRange } from '../address';
import type { Cell } from '../model';
import { touchedCells, type HistoryStep } from './history';

type Captured = [number, number, Cell][];
const cells = (positions: [number, number][]): Captured =>
	positions.map(([row, col]) => [row, col, { value: row * 100 + col }]);
const range: CellRange = { start: { row: 0, col: 0 }, end: { row: 9, col: 9 } };
const step = (...entries: [number, Captured, Captured][]): HistoryStep => ({
	label: 'test',
	structural: false,
	entries: entries.map(([sheet, before, after]) => ({
		before: { kind: 'cells', sheet, ranges: [range], cells: before },
		after: { kind: 'cells', sheet, ranges: [range], cells: after },
	})),
});

/** The list as it was built before: every snapshot through one set per sheet. */
function reference(s: HistoryStep) {
	const out: { sheet: number; row: number; col: number }[] = [];
	const seen = new Map<number, Set<number>>();
	for (const entry of s.entries)
		for (const snapshot of [entry.before, entry.after]) {
			if (snapshot.kind !== 'cells') continue;
			let keys = seen.get(snapshot.sheet);
			if (!keys) seen.set(snapshot.sheet, (keys = new Set()));
			for (const [row, col] of snapshot.cells) {
				if (keys.has(cellKey(row, col))) continue;
				keys.add(cellKey(row, col));
				out.push({ sheet: snapshot.sheet, row, col });
			}
		}
	return out;
}

describe('the cells a step touched', () => {
	it('are listed as before, each once and in the same order', () => {
		const grid = (rows: number, cols: number, skip = 0) =>
			Array.from({ length: rows * cols }, (_, i): [number, number] => [
				Math.floor(i / cols),
				i % cols,
			]).filter((_, i) => !skip || i % skip);
		const steps = [
			step([0, [], cells(grid(10, 10))]),
			step([0, cells(grid(10, 10, 3)), cells(grid(10, 10))]),
			step([0, cells(grid(10, 10)), cells(grid(10, 10, 4))]),
			step([0, cells(grid(4, 4)), []]),
			step([
				0,
				cells([
					[2, 2],
					[0, 0],
					[1, 5],
				]),
				cells([
					[0, 0],
					[3, 1],
				]),
			]),
			step([0, cells(grid(3, 3)), cells(grid(3, 3))], [1, [], cells(grid(2, 2))]),
			step([0, cells(grid(3, 3)), cells(grid(5, 1))], [0, cells(grid(2, 2)), cells(grid(4, 4))]),
		];
		for (const s of steps) expect(touchedCells(s)).toEqual(reference(s));
	});
});
