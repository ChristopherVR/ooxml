import { describe, expect, it } from 'vitest';
import { parseAddress, parseRange } from '../address.js';
import { getCell } from '../cells.js';
import type { CellValue, Workbook, Worksheet } from '../model.js';
import { styleAt } from '../styles.js';
import { createWorkbook } from '../workbook.js';
import { createEditSession } from './session.js';

const A = (ref: string) => {
	const a = parseAddress(ref);
	if (!a) throw new Error(ref);
	return a;
};
const R = (ref: string) => {
	const r = parseRange(ref);
	if (!r) throw new Error(ref);
	return r;
};
const ws = (wb: Workbook): Worksheet => {
	const sheet = wb.sheets[0];
	if (!sheet) throw new Error('sheet');
	return sheet;
};
const values = (wb: Workbook, ref: string): CellValue[] => {
	const r = R(ref);
	const out: CellValue[] = [];
	for (let row = r.start.row; row <= r.end.row; row++)
		for (let col = r.start.col; col <= r.end.col; col++)
			out.push(getCell(ws(wb), row, col)?.value ?? null);
	return out;
};
/** Types `inputs` down column A from A1 and fills the source over `target`. */
const fillDown = (inputs: string[], target: string) => {
	const wb = createWorkbook();
	const s = createEditSession(wb, { recalc: false });
	inputs.forEach((text, i) => s.setCellInput(0, i, 0, text));
	s.fill(0, { start: A('A1'), end: { row: inputs.length - 1, col: 0 } }, R(target));
	return { wb, s };
};

describe('fill series', () => {
	it('continues 1, 2 with 3, 4', () => {
		const { wb } = fillDown(['1', '2'], 'A1:A4');
		expect(values(wb, 'A1:A4')).toEqual([1, 2, 3, 4]);
	});
	it('continues a step of 5 and decimals', () => {
		expect(values(fillDown(['5', '10'], 'A1:A4').wb, 'A3:A4')).toEqual([15, 20]);
		expect(values(fillDown(['0.1', '0.2'], 'A1:A4').wb, 'A3:A4')).toEqual([0.3, 0.4]);
	});
	it('uses the linear trend of three or more numbers', () => {
		expect(values(fillDown(['1', '3', '5'], 'A1:A5').wb, 'A4:A5')).toEqual([7, 9]);
	});
	it('copies a single number', () => {
		expect(values(fillDown(['7'], 'A1:A3').wb, 'A1:A3')).toEqual([7, 7, 7]);
	});
	it('continues weekday names', () => {
		expect(values(fillDown(['Mon'], 'A1:A3').wb, 'A1:A3')).toEqual(['Mon', 'Tue', 'Wed']);
		expect(values(fillDown(['Friday'], 'A1:A4').wb, 'A2:A4')).toEqual([
			'Saturday',
			'Sunday',
			'Monday',
		]);
	});
	it('continues month names and keeps their case', () => {
		expect(values(fillDown(['Jan'], 'A1:A3').wb, 'A2:A3')).toEqual(['Feb', 'Mar']);
		expect(values(fillDown(['NOV'], 'A1:A3').wb, 'A2:A3')).toEqual(['DEC', 'JAN']);
		expect(values(fillDown(['January', 'March'], 'A1:A4').wb, 'A3:A4')).toEqual(['May', 'July']);
	});
	it('continues text with a trailing number', () => {
		expect(values(fillDown(['Item 1'], 'A1:A3').wb, 'A2:A3')).toEqual(['Item 2', 'Item 3']);
		expect(values(fillDown(['Item 1', 'Item 3'], 'A1:A4').wb, 'A3:A4')).toEqual([
			'Item 5',
			'Item 7',
		]);
		expect(values(fillDown(['Row 09'], 'A1:A2').wb, 'A2:A2')).toEqual(['Row 10']);
		expect(values(fillDown(['x01'], 'A1:A2').wb, 'A2:A2')).toEqual(['x02']);
	});
	it('cycles quarters', () => {
		expect(values(fillDown(['Q3'], 'A1:A4').wb, 'A2:A4')).toEqual(['Q4', 'Q1', 'Q2']);
	});
	it('copies plain text and booleans', () => {
		expect(values(fillDown(['abc'], 'A1:A3').wb, 'A1:A3')).toEqual(['abc', 'abc', 'abc']);
		expect(values(fillDown(['TRUE'], 'A1:A2').wb, 'A2:A2')).toEqual([true]);
	});
	it('repeats a mixed pattern', () => {
		expect(values(fillDown(['a', '1'], 'A1:A5').wb, 'A3:A5')).toEqual(['a', 1, 'a']);
	});
	it('increments a single date by one day', () => {
		const { wb } = fillDown(['2024-01-30'], 'A1:A3');
		expect(values(wb, 'A1:A3')).toEqual([45321, 45322, 45323]);
		expect(styleAt(wb, getCell(ws(wb), 2, 0)?.styleId).numFmt).toBe('yyyy-mm-dd');
	});
	it('steps dates by months when the day of month repeats', () => {
		expect(values(fillDown(['2024-01-15', '2024-02-15'], 'A1:A3').wb, 'A3:A3')).toEqual([45366]);
		// Jan 31, Mar 31 -> May 31 (two-month steps, clamped to the month end where needed)
		expect(values(fillDown(['2024-01-31', '2024-03-31'], 'A1:A3').wb, 'A3:A3')).toEqual([45443]);
	});
	it('steps dates by a constant number of days', () => {
		const { wb } = fillDown(['2024-01-01', '2024-01-08'], 'A1:A3');
		expect(values(wb, 'A3:A3')).toEqual([45292 + 14]);
	});
	it('translates formulas relative to the new cell', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb, { recalc: false });
		s.setCellInput(0, 0, 1, '=A1*2');
		s.fill(0, R('B1'), R('B1:B3'));
		expect(getCell(ws(wb), 2, 1)?.formula).toBe('A3*2');
		s.setCellInput(0, 0, 3, '=$A$1+A1');
		s.fill(0, R('D1'), R('D1:F1'));
		expect(getCell(ws(wb), 0, 5)?.formula).toBe('$A$1+C1');
	});
	it('fills right, up and left', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb, { recalc: false });
		s.setCellInput(0, 5, 5, '10');
		s.setCellInput(0, 5, 6, '20');
		s.fill(0, R('F6:G6'), R('F6:I6'));
		expect(values(wb, 'H6:I6')).toEqual([30, 40]);
		s.fill(0, R('F6:G6'), R('D6:G6'));
		expect(values(wb, 'D6:E6')).toEqual([-10, 0]);
		s.setCellInput(0, 9, 0, 'Wed');
		s.fill(0, R('A10'), R('A8:A10'));
		expect(values(wb, 'A8:A9')).toEqual(['Mon', 'Tue']);
	});
	it('fills each lane separately and copies formats', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb, { recalc: false });
		s.setRangeValues(0, A('A1'), [
			[1, 'Mon'],
			[2, 'Tue'],
		]);
		s.applyStyle(0, [R('A1:A2')], { font: { bold: true } });
		s.fill(0, R('A1:B2'), R('A1:B4'));
		expect(values(wb, 'A3:B4')).toEqual([3, 'Wed', 4, 'Thu']);
		expect(styleAt(wb, getCell(ws(wb), 3, 0)?.styleId).font.bold).toBe(true);
	});
	it('accepts a target that only holds the new cells', () => {
		const { wb } = fillDown(['1', '2'], 'A3:A4');
		expect(values(wb, 'A1:A4')).toEqual([1, 2, 3, 4]);
	});
	it('clears target cells where the source is blank', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb, { recalc: false });
		s.setRangeValues(0, A('A1'), [['x'], [null], ['old'], ['old']]);
		s.fill(0, R('A1:A2'), R('A1:A4'));
		expect(values(wb, 'A3:A4')).toEqual(['x', null]);
	});
	it('undoes a fill', () => {
		const { wb, s } = fillDown(['1', '2'], 'A1:A6');
		s.undo();
		expect(values(wb, 'A3:A6')).toEqual([null, null, null, null]);
	});
	it('rejects a target that extends in two directions', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb, { recalc: false });
		expect(() => s.fill(0, R('A1'), R('A1:B2'))).toThrow();
	});
});
