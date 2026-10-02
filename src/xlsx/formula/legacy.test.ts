import { describe, expect, it } from 'vitest';
import { getCell } from '../cells.js';
import { isSpilledCell } from './spill.js';
import { book, engine, get, locate } from './test-helpers.js';

const legacy = (wb: ReturnType<typeof book>, key: string): void => {
	const at = locate(wb, key);
	const cell = getCell(wb.sheets[at.sheet]!, at.row, at.col);
	if (cell) cell.legacyFormula = true;
};

describe('legacy (pre-dynamic-array) formulas', () => {
	it('reduce a range result by implicit intersection instead of spilling', () => {
		const wb = book({ A1: 1, A2: 2, A3: 3, B2: '=A1:A3', C2: '=A1:A3' });
		legacy(wb, 'B2');
		engine(wb).recalculateAll();
		expect(get(wb, 'B2')).toBe(2);
		expect(get(wb, 'C2')).toBe(1);
		const below = getCell(wb.sheets[0]!, 2, 2);
		expect(isSpilledCell(below)).toBe(true);
		expect(isSpilledCell(getCell(wb.sheets[0]!, 2, 1))).toBe(false);
	});
	it('intersect ranges in operators and single-value parameters', () => {
		const wb = book({ A1: -1, A2: -2, A3: -3, B3: '=ABS(A1:A3)', C3: '=A1:A3*10' });
		legacy(wb, 'B3');
		legacy(wb, 'C3');
		engine(wb).recalculateAll();
		expect(get(wb, 'B3')).toBe(3);
		expect(get(wb, 'C3')).toBe(-30);
	});
	it('keep array evaluation inside array parameters (SUMPRODUCT, SUM)', () => {
		const wb = book({
			A1: 1,
			A2: 2,
			A3: 3,
			B1: '=SUMPRODUCT((A1:A3>1)*A1:A3)',
			B2: '=SUM(A1:A3*2)',
			B3: '=INDEX(A1:A3,2)',
		});
		for (const key of ['B1', 'B2', 'B3']) legacy(wb, key);
		engine(wb).recalculateAll();
		expect(get(wb, 'B1')).toBe(5);
		expect(get(wb, 'B2')).toBe(12);
		expect(get(wb, 'B3')).toBe(2);
	});
	it('return the first element of an array result outside the range', () => {
		const wb = book({ D9: '={10,20;30,40}' });
		legacy(wb, 'D9');
		engine(wb).recalculateAll();
		expect(get(wb, 'D9')).toBe(10);
		expect(getCell(wb.sheets[0]!, 8, 4)).toBeUndefined();
	});
	it('change semantics when the flag changes on an incremental recalculation', () => {
		const wb = book({ A1: 1, A2: 2, B1: '=A1:A2' });
		const calc = engine(wb);
		calc.recalculateAll();
		expect(get(wb, 'B2')).toBe(2);
		legacy(wb, 'B1');
		calc.recalculateFrom([locate(wb, 'B1')]);
		expect(get(wb, 'B1')).toBe(1);
		expect(get(wb, 'B2')).toBeNull();
	});
});
