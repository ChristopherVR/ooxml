import { describe, expect, it } from 'vitest';
import { addFuturePrefixes, stripFuturePrefixes } from './formula-text.js';
import { fromStoredSyntax, toStoredSyntax } from './formula-storage.js';

describe('dynamic-array storage syntax', () => {
	it('writes spill references and @ as ANCHORARRAY and SINGLE', () => {
		expect(toStoredSyntax('SUM(A1#)+Sheet2!B2#')).toBe(
			'SUM(ANCHORARRAY(A1))+ANCHORARRAY(Sheet2!B2)',
		);
		expect(toStoredSyntax('@A1:A3+@INDEX(B:B,2)*@Name')).toBe(
			'SINGLE(A1:A3)+SINGLE(INDEX(B:B,2))*SINGLE(Name)',
		);
		expect(toStoredSyntax('Table1[@Qty]&"#@"')).toBe('Table1[@Qty]&"#@"');
		expect(addFuturePrefixes('SUM(A1#)')).toBe('SUM(_xlfn.ANCHORARRAY(A1))');
	});
	it('reads them back', () => {
		expect(stripFuturePrefixes('SUM(_xlfn.ANCHORARRAY(A1))')).toBe('SUM(A1#)');
		expect(stripFuturePrefixes('_xlfn.SINGLE(A1:A3)+1')).toBe('@A1:A3+1');
		expect(fromStoredSyntax('SINGLE(A1+1)')).toBe('SINGLE(A1+1)');
		expect(fromStoredSyntax('ANCHORARRAY(A1:B2)')).toBe('ANCHORARRAY(A1:B2)');
	});
});
