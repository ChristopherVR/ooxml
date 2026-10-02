import { describe, expect, it } from 'vitest';
import { MAX_ROW, parseRange } from '../address.js';
import {
	shiftIndex,
	shiftIndexClamped,
	shiftRange,
	shiftRangeInBand,
	shiftSpan,
	subtractRange,
} from './range-math.js';

const r = (ref: string) => {
	const range = parseRange(ref);
	if (!range) throw new Error(ref);
	return range;
};

describe('shiftSpan', () => {
	it('moves a span when inserting at or before its start', () => {
		expect(shiftSpan(5, 8, { axis: 'row', at: 5, count: 2 })).toEqual({ start: 7, end: 10 });
		expect(shiftSpan(5, 8, { axis: 'row', at: 0, count: 1 })).toEqual({ start: 6, end: 9 });
	});
	it('widens a span when inserting inside it and ignores inserts after it', () => {
		expect(shiftSpan(5, 8, { axis: 'row', at: 6, count: 3 })).toEqual({ start: 5, end: 11 });
		expect(shiftSpan(5, 8, { axis: 'row', at: 9, count: 3 })).toEqual({ start: 5, end: 8 });
	});
	it('trims, moves and removes spans on delete', () => {
		expect(shiftSpan(5, 8, { axis: 'row', at: 0, count: -2 })).toEqual({ start: 3, end: 6 });
		expect(shiftSpan(5, 8, { axis: 'row', at: 6, count: -2 })).toEqual({ start: 5, end: 6 });
		expect(shiftSpan(5, 8, { axis: 'row', at: 4, count: -3 })).toEqual({ start: 4, end: 5 });
		expect(shiftSpan(5, 8, { axis: 'row', at: 5, count: -4 })).toBeUndefined();
		expect(shiftSpan(5, 8, { axis: 'row', at: 2, count: -10 })).toBeUndefined();
	});
	it('keeps whole-column spans reaching the grid edge', () => {
		expect(shiftSpan(0, MAX_ROW, { axis: 'row', at: 3, count: -2 })).toEqual({
			start: 0,
			end: MAX_ROW,
		});
		expect(shiftSpan(0, MAX_ROW, { axis: 'row', at: 3, count: 2 })).toEqual({
			start: 0,
			end: MAX_ROW,
		});
	});
});

describe('index shifting', () => {
	it('shifts and deletes single indices', () => {
		expect(shiftIndex(3, { axis: 'col', at: 2, count: 1 })).toBe(4);
		expect(shiftIndex(1, { axis: 'col', at: 2, count: 1 })).toBe(1);
		expect(shiftIndex(3, { axis: 'col', at: 2, count: -2 })).toBeUndefined();
		expect(shiftIndex(5, { axis: 'col', at: 2, count: -2 })).toBe(3);
	});
	it('drops indices pushed off the grid and clamps deleted ones', () => {
		expect(shiftIndex(MAX_ROW, { axis: 'row', at: 0, count: 1 })).toBeUndefined();
		expect(shiftIndexClamped(3, { axis: 'row', at: 2, count: -4 })).toBe(2);
	});
});

describe('ranges', () => {
	it('shifts a range along one axis', () => {
		expect(shiftRange(r('B2:C3'), { axis: 'row', at: 0, count: 1 })).toEqual(r('B3:C4'));
		expect(shiftRange(r('B2:C3'), { axis: 'col', at: 2, count: 1 })).toEqual(r('B2:D3'));
		expect(shiftRange(r('B2:C3'), { axis: 'col', at: 1, count: -2 })).toBeUndefined();
	});
	it('shifts only ranges entirely inside a band', () => {
		const shift = { axis: 'row' as const, at: 1, count: 1 };
		expect(shiftRangeInBand(r('A3:B3'), shift, 0, 1)).toEqual(r('A4:B4'));
		expect(shiftRangeInBand(r('D3:E3'), shift, 0, 1)).toEqual(r('D3:E3'));
		expect(shiftRangeInBand(r('B3:C3'), shift, 0, 1)).toEqual(r('B3:C3'));
	});
	it('subtracts one rectangle from another', () => {
		expect(subtractRange(r('A1:C3'), r('E5'))).toEqual([r('A1:C3')]);
		expect(subtractRange(r('A1:C3'), r('A1:C3'))).toEqual([]);
		const parts = subtractRange(r('A1:C3'), r('B2'));
		expect(parts).toHaveLength(4);
		const cells = parts.reduce(
			(n, p) => n + (p.end.row - p.start.row + 1) * (p.end.col - p.start.col + 1),
			0,
		);
		expect(cells).toBe(8);
		expect(subtractRange(r('A1:A10'), r('A1:A5'))).toEqual([r('A6:A10')]);
	});
});
