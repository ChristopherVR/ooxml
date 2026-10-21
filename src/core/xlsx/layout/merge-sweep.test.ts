import { describe, expect, it } from 'vitest';
import type { CellRange } from '../address';
import { MergeSweep } from './merge-sweep';

const range = (r1: number, c1: number, r2: number, c2: number): CellRange => ({
	start: { row: r1, col: c1 },
	end: { row: r2, col: c2 },
});

/** What the sweep must answer: the merges covering a cell, read straight from the list. */
function expected(merges: readonly CellRange[], row: number, col: number) {
	const covering = merges.filter(
		(m) => row >= m.start.row && row <= m.end.row && col >= m.start.col && col <= m.end.col,
	);
	return {
		merged: covering.length > 0,
		spanned: covering.some((m) => m.start.row !== m.end.row || m.start.col !== m.end.col),
	};
}

/** A small deterministic generator, so a failure always reproduces. */
function random(seed: number): (n: number) => number {
	let state = seed;
	return (n) => {
		state = (state * 1103515245 + 12345) % 2147483648;
		return state % n;
	};
}

/** Merges that do not overlap, as a valid sheet has them: placed on a grid of free cells. */
function disjointMerges(next: (n: number) => number, count: number): CellRange[] {
	const taken = new Set<string>();
	const merges: CellRange[] = [];
	for (let tries = 0; merges.length < count && tries < count * 50; tries++) {
		const r1 = next(60);
		const c1 = next(12);
		const merge = range(r1, c1, r1 + next(next(2) ? 3 : 25), c1 + next(4));
		const cells: string[] = [];
		for (let r = merge.start.row; r <= merge.end.row; r++)
			for (let c = merge.start.col; c <= merge.end.col; c++) cells.push(`${r}:${c}`);
		if (cells.some((cell) => taken.has(cell))) continue;
		for (const cell of cells) taken.add(cell);
		merges.push(merge);
	}
	return merges;
}

function check(merges: readonly CellRange[], rows: readonly number[]): void {
	const sweep = new MergeSweep(merges);
	for (const row of rows) {
		sweep.moveTo(row);
		for (let col = 0; col < 20; col++) expect(sweep.at(col)).toEqual(expected(merges, row, col));
	}
}

describe('MergeSweep', () => {
	it.each([1, 2, 3, 4, 5, 6])(
		'matches a scan of every merge (disjoint merges, seed %i)',
		(seed) => {
			const next = random(seed);
			const merges = disjointMerges(next, 40);
			expect(merges.length).toBeGreaterThan(20);
			const rows = Array.from({ length: 100 }, (_, row) => row).filter(() => next(3) > 0);
			check(merges, rows);
		},
	);

	it.each([7, 8, 9])('matches a scan of every merge when merges overlap (seed %i)', (seed) => {
		const next = random(seed);
		const merges = Array.from({ length: 30 }, () => {
			const r1 = next(40);
			const c1 = next(15);
			return range(r1, c1, r1 + next(8), c1 + next(next(2) ? 1 : 5));
		});
		check(
			merges,
			Array.from({ length: 60 }, (_, row) => row),
		);
	});

	it('sees a spanning merge around a single-cell one, which only a malformed file has', () => {
		const merges = [range(2, 2, 2, 2), range(1, 1, 4, 3)];
		const sweep = new MergeSweep(merges);
		sweep.moveTo(2);
		expect(sweep.at(2)).toEqual({ merged: true, spanned: true });
		sweep.moveTo(5);
		expect(sweep.at(2)).toEqual({ merged: false, spanned: false });
	});

	it('holds only the merges covering the current row', () => {
		// Tall merges in every other column, ending row by row, plus short ones further down.
		const merges = [
			...Array.from({ length: 200 }, (_, i) => range(0, i * 2, 100 + i, i * 2)),
			...Array.from({ length: 50 }, (_, i) => range(400 + i * 2, 0, 401 + i * 2, 1)),
		];
		const sweep = new MergeSweep(merges);
		for (let row = 0; row < 600; row++) {
			sweep.moveTo(row);
			const covering = merges.filter((m) => row >= m.start.row && row <= m.end.row).length;
			expect(sweep.size).toBe(covering);
		}
	});

	it('skips merges between visited rows and refuses to go up', () => {
		const merges = [range(3, 0, 5, 1), range(20, 0, 20, 0)];
		const sweep = new MergeSweep(merges);
		sweep.moveTo(1);
		expect(sweep.size).toBe(0);
		sweep.moveTo(30);
		expect(sweep.size).toBe(0);
		expect(sweep.at(0)).toEqual({ merged: false, spanned: false });
		expect(() => sweep.moveTo(29)).toThrow(RangeError);
	});
});
