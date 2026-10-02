import { describe, expect, it } from 'vitest';
import { addFuturePrefixes, stripFuturePrefixes, translateReferences } from './formula-text.js';

describe('translateReferences', () => {
	it('moves relative parts and keeps absolute ones', () => {
		expect(translateReferences('B2*C2', 3, 0)).toBe('B5*C5');
		expect(translateReferences('D2/SUM($D$2:$D$6)', 2, 0)).toBe('D4/SUM($D$2:$D$6)');
		expect(translateReferences('$A1+A$1', 1, 1)).toBe('$A2+B$1');
	});

	it('handles sheet-qualified, whole-row and whole-column references', () => {
		expect(translateReferences("Sheet2!A1+'My Sheet'!B2", 1, 1)).toBe("Sheet2!B2+'My Sheet'!C3");
		expect(translateReferences('SUM(A:A)+SUM(1:2)', 1, 1)).toBe('SUM(B:B)+SUM(2:3)');
		expect(translateReferences('SUM($A:B)', 0, 2)).toBe('SUM($A:D)');
	});

	it('leaves strings, function names, structured references and names alone', () => {
		expect(translateReferences('IF(A1="B2",LOG10(A1),Table1[Col A1])', 1, 0)).toBe(
			'IF(A2="B2",LOG10(A2),Table1[Col A1])',
		);
		expect(translateReferences('TaxRate*A1+1.5E+3', 1, 0)).toBe('TaxRate*A2+1.5E+3');
	});

	it('turns references pushed off the grid into #REF!', () => {
		expect(translateReferences('A1', -1, 0)).toBe('#REF!');
	});
});

describe('future function prefixes', () => {
	it('strips and restores _xlfn and _xlws prefixes outside strings', () => {
		const stored = '_xlfn.XLOOKUP(1,A:A,B:B)+_xlfn._xlws.FILTER(A1:A3,B1:B3)&"_xlfn.X("';
		const plain = stripFuturePrefixes(stored);
		expect(plain).toBe('XLOOKUP(1,A:A,B:B)+FILTER(A1:A3,B1:B3)&"_xlfn.X("');
		expect(addFuturePrefixes(plain)).toBe(stored);
		expect(addFuturePrefixes('SUM(A1)+IFS(A1>1,1)')).toBe('SUM(A1)+_xlfn.IFS(A1>1,1)');
	});
});
