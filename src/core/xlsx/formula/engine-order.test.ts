import { describe, expect, it } from 'vitest';
import { getCell } from '../cells.js';
import { createCalcEngine } from './engine.js';
import { book, engine, get, locate, set } from './test-helpers.js';

describe('full recalculation dependency order reuse', () => {
	it('evaluates every formula again, including volatile functions and unchanged spill shapes', () => {
		const wb = book({ A1: 2, D1: '=A1:B1*2', B1: 3, G1: '=E1+RAND()' });
		let calls = 0;
		const e = createCalcEngine(wb, { random: () => ++calls / 10 });
		e.recalculateAll();
		expect(get(wb, 'G1')).toBe(6.1);
		set(wb, 'B1', 5);
		e.recalculateAll();
		expect(get(wb, 'G1')).toBe(10.2);
		set(wb, 'B1', 7);
		e.recalculateAll();
		expect(get(wb, 'G1')).toBe(14.3);
		expect(calls).toBe(3);
	});

	it('reorders a reader inserted before its new spill anchor', () => {
		const wb = book({ G1: '=E1+1', D1: '=A1:B1*2', A1: 2, B1: 3 });
		const e = engine(wb);
		e.recalculateAll();
		expect(get(wb, 'G1')).toBe(7);
		set(wb, 'B1', 5);
		e.recalculateAll();
		expect(get(wb, 'G1')).toBe(11);
		expect(e.circularCells()).toEqual([]);
	});

	it('detects a cycle introduced by a new spill and rebuilds when the spill is removed', () => {
		const wb = book({ A1: '=B1:C1*2' });
		const e = engine(wb);
		e.recalculateAll();
		e.recalculateAll();
		expect(e.circularCells()).toEqual([{ sheet: 0, row: 0, col: 0 }]);
		expect(get(wb, 'A1')).toBe(0);
		expect(get(wb, 'B1')).toBeNull();
		e.recalculateAll();
		expect(e.circularCells()).toEqual([{ sheet: 0, row: 0, col: 0 }]);
		set(wb, 'A1', '=2');
		e.recalculateFrom([locate(wb, 'A1')]);
		e.recalculateAll();
		expect(get(wb, 'A1')).toBe(2);
		expect(e.circularCells()).toEqual([]);
	});

	it('invalidates after incremental spill resizing and structural formula edits', () => {
		const wb = book({ D1: '=SEQUENCE(A1,1,B1)', A1: 2, B1: 3, G1: '=D3+1' });
		const e = engine(wb);
		e.recalculateAll();
		e.recalculateAll();
		set(wb, 'A1', 3);
		e.recalculateFrom([locate(wb, 'A1')]);
		expect(get(wb, 'G1')).toBe(6);
		set(wb, 'B1', 10);
		e.recalculateAll();
		expect(get(wb, 'G1')).toBe(13);
		set(wb, 'D1', '=B1+2');
		e.recalculateFrom([locate(wb, 'D1')]);
		expect(get(wb, 'D3')).toBeNull();
		set(wb, 'G1', '=D1*2');
		e.recalculateFrom([locate(wb, 'G1')]);
		set(wb, 'B1', 20);
		e.recalculateAll();
		expect(get(wb, 'G1')).toBe(44);
		expect(getCell(wb.sheets[0]!, 2, 3)).toBeUndefined();
	});
});
