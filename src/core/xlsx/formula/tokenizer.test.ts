import { describe, expect, it } from 'vitest';
import { FormulaError } from './ast.js';
import { joinTokens, numberLiteral, tokenize } from './tokenizer.js';

const kinds = (formula: string): string[] =>
	tokenize(formula)
		.filter((t) => t.kind !== 'ws')
		.map((t) => `${t.kind}:${t.text}`);

describe('tokenize', () => {
	it('round-trips the exact source text, whitespace included', () => {
		for (const f of [
			'SUM( A1 , B2 )',
			"='My Sheet'!$A$1+1",
			'{1,2;3,4}',
			'Table1[[#This Row],[Amount]]*2',
		]) {
			expect(joinTokens(tokenize(f))).toBe(f);
		}
	});

	it('reads numbers, strings with doubled quotes, booleans and errors', () => {
		expect(kinds('1.5E3+"a""b"&TRUE&#N/A')).toEqual([
			'number:1.5E3',
			'op:+',
			'string:"a""b"',
			'op:&',
			'bool:TRUE',
			'op:&',
			'error:#N/A',
		]);
		expect(tokenize('"a""b"')[0]?.value).toBe('a"b');
	});

	it('distinguishes cell references from function names that look like cells', () => {
		expect(kinds('LOG10(A1)')).toEqual(['func:LOG10', 'open:(', 'ref:A1', 'close:)']);
		expect(kinds('A1B')).toEqual(['name:A1B']);
	});

	it('reads areas, whole columns and whole rows as single tokens', () => {
		expect(kinds('A1:B2+$C:$D+1:3')).toEqual(['ref:A1:B2', 'op:+', 'ref:$C:$D', 'op:+', 'ref:1:3']);
		const cols = tokenize('A:C')[0];
		expect(cols?.ref?.kind).toBe('cols');
		expect(cols?.ref?.end.col).toBe(2);
	});

	it('records absolute markers', () => {
		const ref = tokenize('$A1:B$2')[0]?.ref;
		expect(ref?.start).toEqual({ row: 0, col: 0, rowAbs: false, colAbs: true });
		expect(ref?.end).toEqual({ row: 1, col: 1, rowAbs: true, colAbs: false });
	});

	it('reads sheet prefixes, quoted names with escaped quotes, 3D and external prefixes', () => {
		const quoted = tokenize("'It''s here'!B3")[0];
		expect(quoted?.prefix?.sheet).toBe("It's here");
		expect(quoted?.prefix?.text).toBe("'It''s here'!");
		expect(tokenize('Sheet1:Sheet3!A1')[0]?.prefix).toMatchObject({
			sheet: 'Sheet1',
			sheet2: 'Sheet3',
		});
		expect(tokenize('[1]Data!A1')[0]?.prefix).toMatchObject({ book: '1', sheet: 'Data' });
		const broken = tokenize('Sheet2!#REF!')[0];
		expect(broken?.kind).toBe('ref');
		expect(broken?.ref).toBeUndefined();
	});

	it('reads sheet-qualified names and function prefixes', () => {
		expect(kinds('Sheet1!Rate*_xlfn.XLOOKUP(1,A:A,B:B)')[0]).toBe('name:Sheet1!Rate');
		expect(kinds('_xlfn.XLOOKUP(1)')[0]).toBe('func:_xlfn.XLOOKUP');
	});

	it('reads structured references with nested brackets and escapes', () => {
		expect(kinds("Sales[[#This Row],[Unit '[x']]]")).toEqual([
			"structured:Sales[[#This Row],[Unit '[x']]]",
		]);
		expect(kinds('[@Qty]*2')[0]).toBe('structured:[@Qty]');
	});

	it('reads spill references', () => {
		const token = tokenize('SUM(A1#)')[2];
		expect(token).toMatchObject({ kind: 'ref', text: 'A1#', spill: true });
	});

	it('reads comparison operators greedily', () => {
		expect(kinds('A1<>1')).toContain('op:<>');
		expect(kinds('A1<=1')).toContain('op:<=');
		expect(kinds('A1>=1')).toContain('op:>=');
	});

	it('accepts a leading equals sign', () => {
		expect(kinds('=1+1')).toEqual(['number:1', 'op:+', 'number:1']);
	});

	it('rejects malformed input', () => {
		expect(() => tokenize('"abc')).toThrow(FormulaError);
		expect(() => tokenize('#BAD')).toThrow(FormulaError);
		expect(() => tokenize("'Sheet1!A1")).toThrow(FormulaError);
		expect(() => tokenize('A1 ` 2')).toThrow(FormulaError);
	});

	it('treats an out-of-grid column as a name', () => {
		expect(kinds('XFE1')).toEqual(['name:XFE1']);
		expect(kinds('XFD1')).toEqual(['ref:XFD1']);
	});
});

describe('numberLiteral', () => {
	it('keeps only 15 significant digits like Excel', () => {
		expect(numberLiteral('123456789012345678')).toBe(123456789012345000);
		expect(numberLiteral('0.1234567890123456789')).toBe(0.123456789012345);
		expect(numberLiteral('1.5E3')).toBe(1500);
		expect(numberLiteral('0.000')).toBe(0);
	});
});
