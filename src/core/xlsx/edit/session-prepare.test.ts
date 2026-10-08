// Building the formula graph ahead of the first edit (`prepareCalculation`): in slices, resumable,
// and with edits that arrive before it is done still correct.
import { describe, expect, it } from 'vitest';
import { getCell, putCell } from '../cells';
import type { Cell, Workbook } from '../model';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';

const ROWS = 200;

/** As a file loads it: A = inputs, B = A*10 with stale stored results, C = B+1 saved without. */
function opened(dynamic = false): Workbook {
	const wb = createWorkbook({ sheets: ['Sheet1', 'Sheet2'] });
	const sheet = wb.sheets[0]!;
	const kind: Partial<Cell> = dynamic ? { dynamicArray: true } : { legacyFormula: true };
	for (let r = 0; r < ROWS; r++) {
		putCell(sheet, r, 0, { value: r });
		putCell(sheet, r, 1, { formula: `A${r + 1}*10`, value: -1, ...kind });
		putCell(sheet, r, 2, { formula: `B${r + 1}+1`, value: r === 5 ? null : -2, ...kind });
	}
	putCell(wb.sheets[1]!, 0, 0, { formula: `SUM(Sheet1!B1:B${ROWS})`, value: -3, ...kind });
	return wb;
}

const at = (wb: Workbook, row: number, col: number, sheet = 0) =>
	getCell(wb.sheets[sheet]!, row, col)?.value;

/** An idle deadline that is already spent: every call does one slice of work. */
const spent = () => 0;

describe('prepareCalculation', () => {
	it('builds the graph in resumable slices without evaluating anything', () => {
		const wb = opened();
		const session = createEditSession(wb);
		let slices = 1;
		while (!session.prepareCalculation({ timeRemaining: spent })) slices++;
		expect(slices).toBeGreaterThan(5);
		expect(session.prepareCalculation()).toBe(true);
		expect([at(wb, 3, 1), at(wb, 5, 2), at(wb, 0, 0, 1)]).toEqual([-1, null, -3]);
	});

	it('keeps the stored-values rule for the first edit after preparing', () => {
		const wb = opened();
		const session = createEditSession(wb);
		expect(session.prepareCalculation()).toBe(true);
		session.setCellValue(0, 3, 0, 7);
		expect([at(wb, 3, 1), at(wb, 3, 2)]).toEqual([70, 71]);
		// Not reached by the edit: stored results stand; the one saved without a value is computed.
		expect([at(wb, 4, 1), at(wb, 5, 2)]).toEqual([-1, -1 + 1]);
		expect(at(wb, 0, 0, 1)).toBe(-1 * (ROWS - 1) + 70);
	});

	it('finishes the work when an edit arrives first', () => {
		const wb = opened();
		const session = createEditSession(wb);
		expect(session.prepareCalculation({ timeRemaining: spent })).toBe(false);
		// Row 150 has not been prepared yet; row 0 has.
		session.setCellValue(0, 150, 0, 2);
		session.setCellValue(0, 0, 0, 1);
		expect([at(wb, 150, 1), at(wb, 150, 2), at(wb, 0, 1), at(wb, 0, 2)]).toEqual([20, 21, 10, 11]);
		expect(session.prepareCalculation()).toBe(true);
		session.setCellValue(0, 150, 0, 3);
		expect(at(wb, 150, 2)).toBe(31);
	});

	it('takes formulas typed before the work is done', () => {
		const wb = opened();
		const session = createEditSession(wb);
		session.prepareCalculation({ timeRemaining: spent });
		session.setCellInput(0, 0, 1, '=A1*100');
		session.setCellInput(0, 190, 1, '=A191*100');
		expect([at(wb, 0, 1), at(wb, 0, 2), at(wb, 190, 1), at(wb, 190, 2)]).toEqual([
			0, 1, 19000, 19001,
		]);
		while (!session.prepareCalculation({ timeRemaining: spent }));
		session.setCellValue(0, 190, 0, 1);
		expect(at(wb, 190, 2)).toBe(101);
	});

	it('calculates in full when stored formulas may spill, prepared or not', () => {
		const wb = opened(true);
		const session = createEditSession(wb);
		session.prepareCalculation({ timeRemaining: spent });
		session.setCellValue(0, 0, 0, 1);
		expect([at(wb, 4, 1), at(wb, 4, 2), at(wb, 0, 2)]).toEqual([40, 41, 11]);
	});

	it('is a no-op when edits do not recalculate', () => {
		const manual = opened();
		manual.calcMode = 'manual';
		expect(createEditSession(manual).prepareCalculation({ timeRemaining: spent })).toBe(true);
		const off = createEditSession(opened(), { recalc: false });
		expect(off.prepareCalculation({ timeRemaining: spent })).toBe(true);
	});

	it('rebuilds a prepared graph the engine was not told about', () => {
		const wb = opened();
		const session = createEditSession(wb);
		session.calc.prepare();
		// A formula added behind the engine's back, then the preparation dropped.
		putCell(wb.sheets[0]!, 0, 4, { formula: 'A1*3', value: null, legacyFormula: true });
		session.calc.discardPreparation();
		session.setCellValue(0, 0, 0, 5);
		expect(at(wb, 0, 4)).toBe(15);
	});
});
