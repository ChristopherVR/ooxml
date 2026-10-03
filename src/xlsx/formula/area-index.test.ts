import { describe, expect, it } from 'vitest';
import { type CellRange, rangesIntersect } from '../address.js';
import { AreaIndex } from './area-index.js';

const range = (row: number, col: number, endRow = row, endCol = col): CellRange => ({
	start: { row, col },
	end: { row: endRow, col: endCol },
});

describe('spill footprint indexing', () => {
	it('finds short row footprints by key without scanning neighbouring spills', () => {
		const index = new AreaIndex<number>();
		let rangeReads = 0;
		for (let row = 0; row < 8000; row++) {
			const footprint = range(row, 3, row, 4);
			index.add(
				{
					sheet: 0,
					range: {
						get start() {
							rangeReads++;
							return footprint.start;
						},
						get end() {
							rangeReads++;
							return footprint.end;
						},
					},
				},
				row,
			);
		}
		rangeReads = 0;
		for (let row = 0; row < 8000; row++) {
			const found: number[] = [];
			index.dependents(0, range(row, 4), found);
			expect(found).toEqual([row]);
		}
		expect(rangeReads).toBe(0);
	});

	it('matches range intersections at bucket boundaries and deduplicates multi-cell hits', () => {
		const footprints = [
			range(63, 3, 63, 4),
			range(64, 0, 64, 63),
			range(64, 0, 64, 64),
			range(63, 3, 65, 4),
			range(0, 3, 5000, 4),
		];
		const index = new AreaIndex<number>();
		footprints.forEach((fp, id) => index.add({ sheet: 0, range: fp }, id));
		for (const query of [
			range(63, 3, 64, 4),
			range(64, 63, 64, 64),
			range(0, 0, 5000, 64),
			range(65, 3),
			range(66, 65),
		]) {
			const found: number[] = [];
			index.dependents(0, query, found);
			expect(found.sort()).toEqual(
				footprints.flatMap((fp, id) => (rangesIntersect(fp, query) ? [id] : [])),
			);
			const onOtherSheet = new Set<number>();
			index.dependents(1, query, onOtherSheet);
			expect(onOtherSheet.size).toBe(0);
		}
	});
});
