import { describe, expect, it } from 'vitest';
import { book, calc, calcArray, E, engine, get, set } from './test-helpers.js';

describe('operators and coercion', () => {
	it('evaluates arithmetic with Excel precedence', () => {
		expect(calc('1+2*3')).toBe(7);
		expect(calc('-2^2')).toBe(4);
		expect(calc('2^3^2')).toBe(64);
		expect(calc('(-8)^(1/3)')).toBeCloseTo(-2, 12);
		expect(calc('(-8)^0.5')).toEqual(E.NUM);
		expect(calc('0^0')).toEqual(E.NUM);
	});

	it('treats a blank cell as 0 in arithmetic and "" in text', () => {
		expect(calc('A1+1')).toBe(1);
		expect(calc('A1&"x"')).toBe('x');
		expect(calc('A1')).toBe(0);
		expect(calc('A1=""')).toBe(true);
		expect(calc('A1=0')).toBe(true);
		expect(calc('A1=FALSE')).toBe(true);
	});

	it('coerces numeric text and logicals in arithmetic', () => {
		expect(calc('"1,000"*2')).toBe(2000);
		expect(calc('A1*2', { A1: '3' })).toBe(6);
		expect(calc('TRUE*5')).toBe(5);
		expect(calc('"abc"*2')).toEqual(E.VALUE);
	});

	it('propagates the leftmost error', () => {
		expect(calc('NA()+1/0')).toEqual(E.NA);
		expect(calc('1/0&NA()')).toEqual(E.DIV0);
		expect(calc('A1+1', { A1: '=1/0' })).toEqual(E.DIV0);
	});

	it('snaps a final addition that cancels to zero, like Excel', () => {
		expect(calc('0.1+0.2-0.3')).toBe(0);
		expect(calc('(0.1+0.2-0.3)*1')).not.toBe(0);
		expect(calc('10.1-10')).toBeCloseTo(0.1, 12);
	});

	it('compares at 15 significant digits', () => {
		expect(calc('0.1+0.2=0.3')).toBe(true);
		expect(calc('1+1E-14=1')).toBe(false);
	});

	it('broadcasts arrays elementwise', () => {
		expect(calcArray('{1,2,3}*{10;20}', {}, 2, 3)).toEqual([
			[10, 20, 30],
			[20, 40, 60],
		]);
		expect(calcArray('{1,2,3}+{1,2}', {}, 1, 3)).toEqual([[2, 4, E.NA]]);
		expect(calcArray('-{1,2}%', {}, 1, 2)).toEqual([[-0.01, -0.02]]);
		expect(calcArray('A1:A2>1', { A1: 1, A2: 2 }, 2, 1)).toEqual([[false], [true]]);
	});

	it('lifts scalar functions over ranges and arrays', () => {
		expect(calcArray('ABS({-1,2;-3,4})', {}, 2, 2)).toEqual([
			[1, 2],
			[3, 4],
		]);
		expect(calcArray('LEFT({"abc","de"},{1;2})', {}, 2, 2)).toEqual([
			['a', 'd'],
			['ab', 'de'],
		]);
		expect(calcArray('ROUND(A1:A2,0)', { A1: 1.4, A2: 2.6 }, 2, 1)).toEqual([[1], [3]]);
	});

	it('applies reference operators', () => {
		expect(calc('SUM(A1:B2 B1:C3)', { B1: 5, B2: 7 })).toBe(12);
		expect(calc('A1:A2 C1:C2')).toEqual(E.NULL);
		expect(calc('SUM((A1,B1,A1))', { A1: 1, B1: 2 })).toBe(4);
		expect(calc('AREAS((A1,B1:B3))')).toBe(2);
	});

	it('applies implicit intersection with @', () => {
		const wb = book({ A1: 1, A2: 2, A3: 3, C2: '=@A1:A3*10', D2: '=SINGLE(A1:A3)', E5: '=@A1:A3' });
		engine(wb).recalculateAll();
		expect(get(wb, 'C2')).toBe(20);
		expect(get(wb, 'D2')).toBe(2);
		expect(get(wb, 'E5')).toEqual(E.VALUE);
	});

	it('reports unknown functions and names as #NAME?', () => {
		expect(calc('NOPE(1)')).toEqual(E.NAME);
		expect(calc('undefinedName*2')).toEqual(E.NAME);
	});

	it('reports wrong argument counts as #VALUE!', () => {
		expect(calc('ABS()')).toEqual(E.VALUE);
		expect(calc('ABS(1,2)')).toEqual(E.VALUE);
	});

	it('evaluates stored _xlfn prefixes', () => {
		expect(calc('_xlfn.XLOOKUP(2,{1,2},{"a","b"})')).toBe('b');
		expect(calc('_xlfn.LET(_xlpm.x,5,_xlpm.x*2)')).toBe(10);
	});
});

describe('references', () => {
	it('reads other sheets, with and without quotes', () => {
		const wb = book({ 'Sheet2!B2': 7, 'My Data!A1': 3, A1: "='My Data'!A1+Sheet2!B2" }, [
			'Sheet1',
			'Sheet2',
			'My Data',
		]);
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toBe(10);
	});

	it('returns #REF! for missing sheets and external books', () => {
		expect(calc('Nope!A1')).toEqual(E.REF);
		expect(calc('[1]Sheet1!A1')).toEqual(E.REF);
		expect(calc('#REF!+1')).toEqual(E.REF);
	});

	it('sums a 3D reference across sheets', () => {
		const wb = book(
			{ 'Sheet1!B1': 1, 'Sheet2!B1': 2, 'Sheet3!B1': 3, A1: '=SUM(Sheet1:Sheet3!B1)' },
			['Sheet1', 'Sheet2', 'Sheet3'],
		);
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toBe(6);
	});

	it('spills a multi-cell reference and blank cells become 0', () => {
		expect(calcArray('A1:A3', { A1: 1, A3: 3 }, 3, 1)).toEqual([[1], [0], [3]]);
	});

	it('handles whole columns and rows sparsely', () => {
		expect(calc('SUM(A:A)', { A1: 1, A1000: 2, A1048576: 3 })).toBe(6);
		expect(calc('COUNT(2:2)', { A2: 1, Z2: 2 })).toBe(2);
		expect(calc('COUNTBLANK(B:B)', { B5: 1 })).toBe(1048575);
	});
});

describe('defined names', () => {
	it('resolves global and sheet-local names (local wins)', () => {
		const wb = book({ A1: 2, B1: 5, C1: '=Rate*10', 'Sheet2!C1': '=Rate*10' });
		wb.definedNames.push({ name: 'Rate', formula: 'Sheet1!$A$1' });
		wb.definedNames.push({ name: 'Rate', formula: 'Sheet1!$B$1', localSheet: 1 });
		engine(wb).recalculateAll();
		expect(get(wb, 'C1')).toBe(20);
		expect(get(wb, 'Sheet2!C1')).toBe(50);
	});

	it('evaluates sheet-qualified, constant, nested and self-referencing names', () => {
		const wb = book({ A1: '=Sheet2!Local+Twice', A2: '=Loop', A3: '=SUM(Data)', B1: 1, B2: 2 });
		wb.definedNames.push({ name: 'Local', formula: '100', localSheet: 1 });
		wb.definedNames.push({ name: 'Base', formula: '21' });
		wb.definedNames.push({ name: 'Twice', formula: 'Base*2' });
		wb.definedNames.push({ name: 'Loop', formula: 'Loop+1' });
		wb.definedNames.push({ name: 'Data', formula: 'Sheet1!$B$1:$B$2' });
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toBe(142);
		expect(get(wb, 'A2')).toEqual(E.NAME);
		expect(get(wb, 'A3')).toBe(3);
	});

	it('calls LAMBDA functions stored in names', () => {
		const wb = book({ A1: '=Hyp(3,4)' });
		wb.definedNames.push({ name: 'Hyp', formula: 'LAMBDA(a,b,SQRT(a^2+b^2))' });
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toBe(5);
	});
});

describe('LET and LAMBDA', () => {
	it('binds names in order', () => {
		expect(calc('LET(a,1,b,a+1,c,b*10,c)')).toBe(20);
		expect(calc('LET(x,{1,2,3},SUM(x*2))')).toBe(12);
		expect(calc('LET(f,LAMBDA(n,n*n),f(4)+f(2))')).toBe(20);
		expect(calc('LET(a,1)')).toEqual(E.VALUE);
	});

	it('returns #CALC! for an uncalled LAMBDA and #VALUE! for too many arguments', () => {
		expect(calc('LAMBDA(x,x)')).toEqual(E.CALC);
		expect(calc('LAMBDA(x,x)(1,2)')).toEqual(E.VALUE);
	});
});

describe('structured references', () => {
	function tableBook() {
		const wb = book({
			A1: 'Item',
			B1: 'Qty',
			C1: 'Price',
			D1: 'Total',
			A2: 'pen',
			B2: 2,
			C2: 1.5,
			D2: '=[@Qty]*[@Price]',
			A3: 'ink',
			B3: 3,
			C3: 4,
			D3: '=Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]',
			A4: 'Total',
			B4: '=SUBTOTAL(109,Sales[Qty])',
			F1: '=SUM(Sales[Total])',
			F2: '=ROWS(Sales[#All])',
			F3: '=INDEX(Sales[#Headers],3)',
			F4: '=Sales[[#Totals],[Qty]]',
			F5: '=COLUMNS(Sales[[Qty]:[Total]])',
			F6: '=ROWS(Sales)',
			F7: '=Sales[Nope]',
			F8: '=Other[Qty]',
		});
		wb.sheets[0]?.tables.push({
			id: 1,
			name: 'Sales',
			displayName: 'Sales',
			range: { start: { row: 0, col: 0 }, end: { row: 3, col: 3 } },
			headerRow: true,
			totalsRow: true,
			columns: [{ name: 'Item' }, { name: 'Qty' }, { name: 'Price' }, { name: 'Total' }],
		});
		return wb;
	}

	it('resolves columns, this-row items and special items', () => {
		const wb = tableBook();
		engine(wb).recalculateAll();
		expect(get(wb, 'D2')).toBe(3);
		expect(get(wb, 'D3')).toBe(12);
		expect(get(wb, 'B4')).toBe(5);
		expect(get(wb, 'F1')).toBe(15);
		expect(get(wb, 'F2')).toBe(4);
		expect(get(wb, 'F3')).toBe('Price');
		expect(get(wb, 'F4')).toBe(5);
		expect(get(wb, 'F5')).toBe(3);
		expect(get(wb, 'F6')).toBe(2);
		expect(get(wb, 'F7')).toEqual(E.REF);
		expect(get(wb, 'F8')).toEqual(E.NAME);
	});
});

describe('1904 date system', () => {
	it('uses the workbook epoch', () => {
		const wb = book({ A1: '=DATE(1904,1,2)', A2: '=YEAR(0)', A3: '="1904-01-03"+0' });
		wb.date1904 = true;
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toBe(1);
		expect(get(wb, 'A2')).toBe(1904);
		expect(get(wb, 'A3')).toBe(2);
	});
});

describe('volatile functions', () => {
	it('uses the injected clock and random source', () => {
		const wb = book({ A1: '=TODAY()', A2: '=NOW()', A3: '=RAND()', A4: '=RANDBETWEEN(1,6)' });
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toBe(46298);
		expect(get(wb, 'A2')).toBeCloseTo(46298.5, 8);
		const r = get(wb, 'A3') as number;
		expect(r).toBeGreaterThanOrEqual(0);
		expect(r).toBeLessThan(1);
		expect([1, 2, 3, 4, 5, 6]).toContain(get(wb, 'A4'));
	});

	it('recomputes volatile cells on every incremental recalculation', () => {
		const wb = book({ A1: '=RAND()', B1: 1 });
		const calc = engine(wb);
		calc.recalculateAll();
		const first = get(wb, 'A1');
		set(wb, 'B1', 2);
		calc.recalculateFrom([{ sheet: 0, row: 0, col: 1 }]);
		expect(get(wb, 'A1')).not.toBe(first);
	});
});
