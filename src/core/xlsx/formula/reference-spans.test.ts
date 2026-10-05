import { describe, expect, it } from 'vitest';
import { MAX_COL, MAX_ROW } from '../address.js';
import { referenceSpans } from './transform.js';

describe('referenceSpans', () => {
	it('gives offsets of cell, area and qualified references', () => {
		const f = "SUM(A1:B2)+'My Sheet'!$C$3*Sheet2!D4";
		const spans = referenceSpans(f, 'Sheet1');
		expect(spans.map((s) => f.slice(s.start, s.end))).toEqual([
			'A1:B2',
			"'My Sheet'!$C$3",
			'Sheet2!D4',
		]);
		expect(spans.map((s) => s.text)).toEqual(spans.map((s) => f.slice(s.start, s.end)));
		expect(spans[0]).toMatchObject({
			sheet: 'Sheet1',
			range: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
		});
		expect(spans[1]?.sheet).toBe('My Sheet');
		expect(spans[2]?.range).toEqual({ start: { row: 3, col: 3 }, end: { row: 3, col: 3 } });
	});

	it('counts a leading = in the offsets', () => {
		const spans = referenceSpans('=A1+1', 'S');
		expect(spans).toHaveLength(1);
		expect(spans[0]).toMatchObject({ start: 1, end: 3, text: 'A1' });
	});

	it('covers whole rows and columns', () => {
		const spans = referenceSpans('SUM(A:B)+SUM(2:3)', 'S');
		expect(spans.map((s) => s.text)).toEqual(['A:B', '2:3']);
		expect(spans[0]?.range).toEqual({ start: { row: 0, col: 0 }, end: { row: MAX_ROW, col: 1 } });
		expect(spans[1]?.range).toEqual({ start: { row: 1, col: 0 }, end: { row: 2, col: MAX_COL } });
	});

	it('skips strings, function names, #REF!, 3D and external references', () => {
		const f = 'IF(A1="B2",LOG10(C3),#REF!)+[1]Sheet1!A1+Sheet1:Sheet3!A1+Sheet2!#REF!';
		expect(referenceSpans(f, 'S').map((s) => s.text)).toEqual(['A1', 'C3']);
	});

	it('returns nothing for an unparsable formula', () => {
		expect(() => referenceSpans('SUM(A1', 'S')).not.toThrow();
		expect(referenceSpans('"unterminated', 'S')).toEqual([]);
	});
});
