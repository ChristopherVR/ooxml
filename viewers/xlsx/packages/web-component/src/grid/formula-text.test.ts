import { describe, expect, it } from 'vitest';
import {
	acceptsReference,
	activeParam,
	callAtCaret,
	coloredReferences,
	functionMatches,
	nameAtCaret,
	referenceTarget,
	syntaxParams,
	toggleAbsolute,
} from './formula-text.js';

describe('formula text helpers', () => {
	it('colours references outside strings and skips function names', () => {
		const refs = coloredReferences(
			'=SUM(A1:B2,$C$3)+LOG10(4)&"D5"+Sheet2!E6+\'My Sheet\'!F7+A:A+3:4',
		);
		expect(refs.map((r) => r.text)).toEqual([
			'A1:B2',
			'$C$3',
			'Sheet2!E6',
			"'My Sheet'!F7",
			'A:A',
			'3:4',
		]);
		expect(refs[0]?.start).toBe(5);
	});

	it('colours references, same text sharing a colour', () => {
		const refs = coloredReferences('=A1+B2+a1');
		expect(refs.map((r) => r.index)).toEqual([0, 1, 0]);
		expect(refs[0]?.start).toBe(1);
		expect(refs[0]?.color).toBe(refs[2]?.color);
		expect(coloredReferences('A1+B2')).toEqual([]);
	});

	it('knows when a pointed reference may be inserted', () => {
		expect(acceptsReference('=', 1)).toBe(true);
		expect(acceptsReference('=SUM(', 5)).toBe(true);
		expect(acceptsReference('=A1+', 4)).toBe(true);
		expect(acceptsReference('=A1', 3)).toBe(false);
		expect(acceptsReference('="a,', 4)).toBe(false);
		expect(acceptsReference('abc', 3)).toBe(false);
	});

	it('finds the function name being typed and matches the catalog', () => {
		expect(nameAtCaret('=SU', 3)).toEqual({ start: 1, prefix: 'SU' });
		expect(nameAtCaret('=A1+VLO', 7)).toEqual({ start: 4, prefix: 'VLO' });
		expect(nameAtCaret('SU', 2)).toBeUndefined();
		const catalog = [
			{ name: 'SUM', category: 'Math', syntax: 'SUM(number1, [number2], ...)', description: '' },
			{
				name: 'SUMIF',
				category: 'Math',
				syntax: 'SUMIF(range, criteria, [sum_range])',
				description: '',
			},
			{ name: 'AVERAGE', category: 'Stat', syntax: 'AVERAGE(number1, ...)', description: '' },
		];
		expect(functionMatches('su', 12, catalog).map((f) => f.name)).toEqual(['SUM', 'SUMIF']);
	});

	it('finds the call and argument at the caret', () => {
		expect(callAtCaret('=IF(A1>0,', 9)).toEqual({ name: 'IF', argIndex: 1 });
		expect(callAtCaret('=SUM(MAX(1,2),', 14)).toEqual({ name: 'SUM', argIndex: 1 });
		expect(callAtCaret('=SUM(1)', 7)).toBeUndefined();
		const { params } = syntaxParams('SUM(number1, [number2], ...)');
		expect(params).toEqual(['number1', '[number2]', '...']);
		expect(activeParam(params, 5)).toBe(1);
	});

	it('resolves reference targets and cycles absolute references with F4', () => {
		expect(referenceTarget("'My Sheet'!$B$2:C3")).toEqual({
			sheet: 'My Sheet',
			range: { start: { row: 1, col: 1 }, end: { row: 2, col: 2 } },
		});
		expect(referenceTarget('C:A')?.range.start.col).toBe(0);
		const once = toggleAbsolute('=A1+1', 2)!;
		expect(once.text).toBe('=$A$1+1');
		expect(toggleAbsolute(once.text, 2)!.text).toBe('=A$1+1');
		expect(toggleAbsolute('=1+2', 2)).toBeUndefined();
	});
});
