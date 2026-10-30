import { describe, expect, it } from 'vitest';
import { parseBreakPositions, splitDelimited, splitFixedWidth } from './text-columns';

describe('splitDelimited', () => {
	it('honours a text qualifier and doubled qualifiers', () => {
		expect(
			splitDelimited('"a,b",c', { delimiters: [','], consecutive: false, qualifier: '"' }),
		).toEqual(['a,b', 'c']);
		expect(
			splitDelimited('"say ""hi""",x', { delimiters: [','], consecutive: false, qualifier: '"' }),
		).toEqual(['say "hi"', 'x']);
	});

	it('treats consecutive delimiters as one when asked', () => {
		expect(splitDelimited('a  b', { delimiters: [' '], consecutive: true, qualifier: '' })).toEqual(
			['a', 'b'],
		);
		expect(
			splitDelimited('a  b', { delimiters: [' '], consecutive: false, qualifier: '' }),
		).toEqual(['a', '', 'b']);
	});
});

describe('splitFixedWidth', () => {
	it('ends a column after each break position (1-based)', () => {
		expect(splitFixedWidth('ABCDEFGHIJ', [5])).toEqual(['ABCDE', 'FGHIJ']);
		expect(splitFixedWidth('ABCDEFGHIJ', [2, 5])).toEqual(['AB', 'CDE', 'FGHIJ']);
	});

	it('keeps the column count for a short line and returns the whole line without breaks', () => {
		expect(splitFixedWidth('ABC', [5, 9])).toEqual(['ABC', '', '']);
		expect(splitFixedWidth('  pad', [])).toEqual(['  pad']);
	});

	it('sorts and de-duplicates the breaks and keeps code points whole', () => {
		expect(splitFixedWidth('abcdef', [4, 2, 4])).toEqual(['ab', 'cd', 'ef']);
		expect(splitFixedWidth('😀ab', [1])).toEqual(['😀', 'ab']);
	});
});

describe('parseBreakPositions', () => {
	it('reads comma-separated positions, ascending and unique', () => {
		expect(parseBreakPositions('12, 5,5')).toEqual([5, 12]);
		expect(parseBreakPositions(' 3 ')).toEqual([3]);
		expect(parseBreakPositions('')).toEqual([]);
		expect(parseBreakPositions(' , ')).toEqual([]);
	});

	it('refuses anything that is not a whole number from 1', () => {
		for (const bad of ['0', '-1', '2.5', 'a', '5 6', '1e3', '3;4'])
			expect(parseBreakPositions(bad), bad).toBeUndefined();
	});
});
