// Dynamic references (OFFSET, INDEX, INDIRECT, `:` between computed references), on-demand
// computation depth and deeply nested formulas.
import { describe, expect, it } from 'vitest';
import { putCell } from '../cells.js';
import { book, engine, get, locate, set } from './test-helpers.js';

const at = (wb: Parameters<typeof locate>[0], key: string) => locate(wb, key);

describe('reference arguments are not static precedents', () => {
	it('reads a neighbour through OFFSET or INDEX of its own cell without a circular reference', () => {
		const wb = book({
			B1: 1,
			B2: '=OFFSET(B2,-1,0)+1',
			C1: 1,
			C2: '=INDEX(C:C,ROW()-1)+1',
			D1: 1,
			D2: '=INDIRECT("D"&ROW()-1)+1',
			E1: 1,
			E2: '=SUM(INDEX(E:E,1):INDEX(E:E,ROW()-1))+1',
			F5: '=ROW(F5)+COLUMNS(F:F)',
		});
		const e = engine(wb);
		e.recalculateAll();
		expect(['B2', 'C2', 'D2', 'E2', 'F5'].map((k) => get(wb, k))).toEqual([2, 2, 2, 2, 6]);
		expect(wb.warnings).toEqual([]);
		expect(e.circularCells()).toEqual([]);
	});

	it('still reports a loop found while evaluating', () => {
		const wb = book({ A1: '=OFFSET(A1,0,0)+1', B1: '=INDEX(B1:B2,1)+1' });
		const e = engine(wb);
		e.recalculateAll();
		expect([get(wb, 'A1'), get(wb, 'B1')]).toEqual([0, 0]);
		expect(e.circularCells().length).toBe(2);
		expect(wb.warnings.length).toBe(1);
	});

	it('recalculates INDEX through its array argument when a cell inside changes', () => {
		const wb = book({ A1: 1, A2: 2, A3: 3, B1: '=INDEX(A1:A3,2)*10' });
		const e = engine(wb);
		e.recalculateAll();
		set(wb, 'A2', 7);
		e.recalculateFrom([at(wb, 'A2')]);
		expect(get(wb, 'B1')).toBe(70);
	});
});

describe('ranges between computed references', () => {
	it('recalculates incrementally like a full recalculation', () => {
		const wb = book({ A1: 1, A2: 2, A3: 3, B1: 10, B2: 20, B3: 30, C1: 100, C2: 200, C3: 300 });
		wb.definedNames.push({ name: 'Start', formula: 'Sheet1!$A$1' });
		wb.definedNames.push({ name: 'Stop', formula: 'Sheet1!$A$3' });
		set(wb, 'E1', '=SUM(Start:Stop)');
		set(wb, 'E2', '=SUM(A1:INDEX(C:C,3))');
		set(wb, 'E3', '=SUM(A1:CHOOSE(1,C3))');
		set(wb, 'E4', '=SUM(INDEX(A:A,1):INDEX(A:A,3))');
		const keys = ['E1', 'E2', 'E3', 'E4'];
		const e = engine(wb);
		e.recalculateAll();
		expect(keys.map((k) => get(wb, k))).toEqual([6, 666, 666, 6]);
		set(wb, 'A2', 1000);
		set(wb, 'B2', 5000);
		e.recalculateFrom([at(wb, 'A2'), at(wb, 'B2')]);
		const incremental = keys.map((k) => get(wb, k));
		e.recalculateAll();
		expect(incremental).toEqual(keys.map((k) => get(wb, k)));
		expect(incremental).toEqual([1004, 6644, 6644, 1004]);
	});
});

describe('formulas computed on demand', () => {
	it('evaluates a 10000-long INDIRECT chain without deep recursion', () => {
		const n = 10_000;
		const wb = book({});
		const ws = wb.sheets[0];
		if (!ws) throw new Error('no sheet');
		for (let r = 0; r < n - 1; r++)
			putCell(ws, r, 0, { value: null, formula: 'INDIRECT("A"&(ROW()+1))+1' });
		putCell(ws, n - 1, 0, { value: 0 });
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toBe(n - 1);
	});

	it('evaluates a long OFFSET chain that reads its own column', () => {
		const n = 3000;
		const wb = book({});
		const ws = wb.sheets[0];
		if (!ws) throw new Error('no sheet');
		for (let r = 0; r < n - 1; r++)
			putCell(ws, r, 1, { value: null, formula: 'OFFSET(B1,ROW(),0)+1' });
		putCell(ws, n - 1, 1, { value: 5 });
		engine(wb).recalculateAll();
		expect(get(wb, 'B1')).toBe(n + 4);
		expect(wb.warnings).toEqual([]);
	});

	it('ends a long loop of INDIRECT reads with a circular reference', () => {
		const n = 500;
		const wb = book({});
		const ws = wb.sheets[0];
		if (!ws) throw new Error('no sheet');
		for (let r = 0; r < n; r++) {
			const next = r === n - 1 ? 1 : r + 2;
			putCell(ws, r, 0, { value: null, formula: `INDIRECT("A${next}")+1` });
		}
		const e = engine(wb);
		expect(() => e.recalculateAll()).not.toThrow();
		expect(e.circularCells().length).toBeGreaterThan(0);
		expect(wb.warnings.length).toBe(1);
	});
});

describe('deeply nested formulas', () => {
	const evaluate = (formula: string) => {
		const wb = book({});
		set(wb, 'A1', `=${formula}`);
		engine(wb).recalculateAll();
		return get(wb, 'A1');
	};

	it('evaluates a 4000-term operator chain', () => {
		expect(evaluate(Array(4000).fill('1').join('+'))).toBe(4000);
		expect(evaluate(Array(2000).fill('2').join('*0.5*'))).toBe(2);
	});

	it('shows #VALUE! instead of throwing for nesting past the limits', () => {
		expect(evaluate(`${'('.repeat(4000)}1${')'.repeat(4000)}`)).toEqual({ error: '#VALUE!' });
		expect(evaluate(`${'-'.repeat(8000)}1`)).toEqual({ error: '#VALUE!' });
		expect(evaluate(`${'IF(1,'.repeat(3000)}1${')'.repeat(3000)}`)).toEqual({ error: '#VALUE!' });
	});

	it('allows 64 nested functions, like Excel, and rejects 65', () => {
		expect(evaluate(`${'ABS('.repeat(64)}-1${')'.repeat(64)}`)).toBe(1);
		expect(evaluate(`${'ABS('.repeat(65)}-1${')'.repeat(65)}`)).toEqual({ error: '#VALUE!' });
	});
});
