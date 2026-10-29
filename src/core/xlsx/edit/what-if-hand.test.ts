// What-If regressions whose expected values are worked out by hand, not read from Excel. They
// guard the arithmetic, not Excel parity: results have not been compared with real Excel output.
import { describe, expect, it } from 'vitest';
import { getCell } from '../cells';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';

const book = () => {
	const wb = createWorkbook({ sheets: ['Sheet1'] });
	return { wb, s: createEditSession(wb) };
};
const at = (wb: ReturnType<typeof book>['wb'], row: number, col: number) =>
	getCell(wb.sheets[0]!, row, col)?.value;

describe('Goal Seek by hand', () => {
	it('solves 2x + 1 = 11 for x = 5', () => {
		const { wb, s } = book();
		s.setCellValue(0, 0, 0, 0);
		s.setCellInput(0, 1, 0, '=2*A1+1');
		const r = s.goalSeek({
			sheet: 0,
			target: { row: 1, col: 0 },
			goal: 11,
			changing: { row: 0, col: 0 },
		});
		expect(r.found).toBe(true);
		expect(at(wb, 0, 0) as number).toBeCloseTo(5, 3);
	});

	it('finds the positive root of x^2 = 49 from a positive start', () => {
		const { wb, s } = book();
		s.setCellValue(0, 0, 0, 3);
		s.setCellInput(0, 1, 0, '=A1^2');
		const r = s.goalSeek({
			sheet: 0,
			target: { row: 1, col: 0 },
			goal: 49,
			changing: { row: 0, col: 0 },
		});
		expect(r.found).toBe(true);
		expect(at(wb, 0, 0) as number).toBeCloseTo(7, 2);
	});
});

describe('data tables by hand', () => {
	it('computes a row-oriented one-variable table (values across the top)', () => {
		const { wb, s } = book();
		s.setCellValue(0, 0, 0, 2);
		[1, 2, 3].forEach((v, i) => s.setCellValue(0, 3, 1 + i, v));
		s.setCellInput(0, 4, 0, '=A1+100');
		s.setCellInput(0, 5, 0, '=A1*A1');
		s.createDataTable(0, {
			range: { start: { row: 3, col: 0 }, end: { row: 5, col: 3 } },
			rowInput: 'A1',
		});
		expect([1, 2, 3].map((i) => at(wb, 4, i))).toEqual([101, 102, 103]);
		expect([1, 2, 3].map((i) => at(wb, 5, i))).toEqual([1, 4, 9]);
		expect(at(wb, 0, 0)).toBe(2);
	});
});

describe('scenarios by hand', () => {
	it('shows a scenario, recalculates dependents and undoes in one step', () => {
		const { wb, s } = book();
		s.setCellValue(0, 0, 0, 5);
		s.setCellInput(0, 0, 1, '=A1+1');
		const big = s.addScenario(0, { name: 'Big', cells: 'A1', values: { A1: '10' } });
		s.showScenario(0, big);
		expect(at(wb, 0, 0)).toBe(10);
		expect(at(wb, 0, 1)).toBe(11);
		s.undo();
		expect(at(wb, 0, 0)).toBe(5);
		expect(at(wb, 0, 1)).toBe(6);
	});
});
