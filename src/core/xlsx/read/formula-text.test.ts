import { describe, expect, it } from 'vitest';
import { addFuturePrefixes, stripFuturePrefixes, translateReferences } from './formula-text';

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

	it('stores LET and LAMBDA names with the _xlpm prefix Excel requires, in scope only', () => {
		const cases: [string, string][] = [
			['LET(x,2,x*3)', '_xlfn.LET(_xlpm.x,2,_xlpm.x*3)'],
			[
				'REDUCE(0,SEQUENCE(3),LAMBDA(a,b,a+b))',
				'_xlfn.REDUCE(0,_xlfn.SEQUENCE(3),_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b))',
			],
			[
				'LET(f,LAMBDA(f,n,IF(n=0,0,1+f(f,n-1))),f(f,10))',
				'_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.f,_xlpm.n,IF(_xlpm.n=0,0,1+_xlpm.f(_xlpm.f,_xlpm.n-1))),_xlpm.f(_xlpm.f,10))',
			],
			// A name outside the LET that declares it, and text, are left alone.
			['SUM(x,LET(x,1,x))&"x"', 'SUM(x,_xlfn.LET(_xlpm.x,1,_xlpm.x))&"x"'],
			['LAMBDA(x,x+1)(5)', '_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(5)'],
		];
		for (const [plain, stored] of cases) {
			expect(addFuturePrefixes(plain)).toBe(stored);
			expect(stripFuturePrefixes(stored)).toBe(plain);
		}
		expect(addFuturePrefixes('_xlfn.LET(_xlpm.x,5,_xlpm.x*2)')).toBe(
			'_xlfn.LET(_xlpm.x,5,_xlpm.x*2)',
		);
	});
});
