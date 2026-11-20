import { describe, expect, it } from 'vitest';
import { getCell, putCell } from '../cells';
import type { Workbook } from '../model';
import { createWorkbook } from '../workbook';
import { calculateOnOpen } from './calc-mode';
import { applyCalculatedCells, calculateOnOpenCells } from './calc-results';

/** The same file read twice: what a worker and the page each hold. */
function workbook(): Workbook {
	const wb = createWorkbook({ sheets: ['S', 'T'] });
	const s = wb.sheets[0]!;
	putCell(s, 0, 0, { value: 2 });
	putCell(s, 1, 0, { value: 5 });
	putCell(s, 0, 1, { value: 99, formula: 'A1*3' });
	putCell(s, 0, 2, { value: 0, formula: 'SEQUENCE(3)' });
	putCell(s, 0, 3, {
		value: 0,
		formula: 'A1:A2*2',
		arrayRange: { start: { row: 0, col: 3 }, end: { row: 1, col: 3 } },
	});
	// A spill the file stored for a formula that no longer spills.
	putCell(s, 4, 0, { value: 0, formula: 'A1+1' });
	putCell(s, 5, 0, { value: 7, spillAnchor: { row: 4, col: 0 } } as never);
	putCell(s, 6, 0, { value: 8, spillAnchor: { row: 4, col: 0 }, styleId: 3 } as never);
	putCell(s, 0, 4, { value: 0, formula: 'E1+1' });
	putCell(wb.sheets[1]!, 2, 2, { value: 'old', formula: 'S!B1&"x"' });
	wb.fullCalcOnLoad = true;
	return wb;
}

const cellsOf = (wb: Workbook) =>
	wb.sheets.map((sheet) => [...sheet.rows].map(([row, line]) => [row, [...line]]));

describe('calculation results brought back from another copy', () => {
	it('leaves a copy exactly as calculating it in place would', () => {
		const inPlace = workbook();
		calculateOnOpen(inPlace);
		const worker = workbook();
		const { calculation, cells } = calculateOnOpenCells(worker);
		expect(calculation).toEqual({ ran: true });
		const page = workbook();
		applyCalculatedCells(page, structuredClone(cells!));
		expect(cellsOf(page)).toEqual(cellsOf(inPlace));
		expect(page.warnings).toEqual(inPlace.warnings);
		expect(getCell(page.sheets[0]!, 2, 2)).toEqual({ value: 3, spillAnchor: { row: 0, col: 2 } });
		expect(getCell(page.sheets[0]!, 5, 0)).toBeUndefined();
		expect(getCell(page.sheets[1]!, 2, 2)?.value).toBe('6x');
	});

	it('brings nothing back when the calculation did not run or stopped', () => {
		const manual = workbook();
		manual.calcMode = 'manual';
		expect(calculateOnOpenCells(manual)).toEqual({ calculation: { ran: false } });
		const slow = workbook();
		for (let row = 10; row < 200; row++)
			putCell(slow.sheets[0]!, row, 0, { value: 0, formula: 'ROW()' });
		expect(calculateOnOpenCells(slow, { timeLimitMs: 0 })).toEqual({
			calculation: { ran: true, stopped: 'time' },
		});
	});
});
