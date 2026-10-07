import { describe, expect, it } from 'vitest';
import { MAX_COL, MAX_ROW } from './address.js';
import { parseR1C1Range } from './address-r1c1.js';

describe('parseR1C1Range', () => {
	it('normalizes relative cell ranges and case', () => {
		expect(parseR1C1Range('r[+1]c[1]:r[-1]c[-1]', { row: 1, col: 1 })).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 2, col: 2 },
		});
	});

	it('normalizes whole-axis ranges without materializing their cells', () => {
		expect(parseR1C1Range('R3:R1', { row: 0, col: 0 })).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 2, col: MAX_COL },
		});
		expect(parseR1C1Range('C3:C1', { row: 0, col: 0 })).toEqual({
			start: { row: 0, col: 0 },
			end: { row: MAX_ROW, col: 2 },
		});
	});

	it('wraps relative coordinates independently in each dimension', () => {
		expect(parseR1C1Range('R[-1]C[1]', { row: 0, col: MAX_COL })).toEqual({
			start: { row: MAX_ROW, col: 0 },
			end: { row: MAX_ROW, col: 0 },
		});
	});

	it('rejects invalid base addresses and malformed endpoints', () => {
		for (const base of [
			{ row: NaN, col: 0 },
			{ row: 0.5, col: 0 },
			{ row: 0, col: MAX_COL + 1 },
		])
			expect(parseR1C1Range('RC', base)).toBeUndefined();
		for (const text of ['R1:C1', 'R1C1:R2', 'RC:RC:RC', 'R[]C', 'C[16384]', 'R1048577C1'])
			expect(parseR1C1Range(text, { row: 0, col: 0 })).toBeUndefined();
	});
});
