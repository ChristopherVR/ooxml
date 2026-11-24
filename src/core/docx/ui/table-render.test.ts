import { describe, expect, it } from 'vitest';
import { tableGridColumns } from './table-render';

describe('tableGridColumns', () => {
	it('gives the grid widths in pixels when there is one per column (#322)', () => {
		expect(tableGridColumns({ grid: JSON.stringify([2340, 2100, 2100, 2100]) }, 4)).toEqual([
			156, 140, 140, 140,
		]);
	});

	it('ignores a missing, malformed or mismatched grid', () => {
		expect(tableGridColumns({}, 2)).toBeNull();
		expect(tableGridColumns({ grid: '[1440' }, 1)).toBeNull();
		expect(tableGridColumns({ grid: JSON.stringify([1440, 1440]) }, 3)).toBeNull();
		expect(tableGridColumns({ grid: JSON.stringify([1440, 0]) }, 2)).toBeNull();
	});
});
