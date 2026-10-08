import { describe, expect, it } from 'vitest';
import { getCell } from '../cells';
import type { Cell, Workbook } from '../model';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';

/** A workbook as a file loads it: formulas with the results Excel stored (some stale). */
function opened(dynamic = false): Workbook {
	const wb = createWorkbook({ sheets: ['Sheet1'] });
	const sheet = wb.sheets[0]!;
	const put = (row: number, col: number, cell: Cell) => {
		let line = sheet.rows.get(row);
		if (!line) sheet.rows.set(row, (line = new Map()));
		line.set(col, cell);
	};
	const formula = (text: string, value: Cell['value']): Cell =>
		dynamic
			? { formula: text, value, dynamicArray: true }
			: { formula: text, value, legacyFormula: true };
	put(0, 0, { value: 2 });
	put(0, 1, formula('A1*10', 20));
	// Stored results that disagree with the inputs: Excel shows them until something they read
	// changes, and so does the editor.
	put(1, 0, { value: 3 });
	put(1, 1, formula('A2*10', 999));
	put(2, 1, formula('B2+1', 1000));
	// Saved without a result (a generator that does not calculate).
	put(3, 1, formula('A1+A2', null));
	return wb;
}

const value = (wb: Workbook, row: number, col: number) => getCell(wb.sheets[0]!, row, col)?.value;

describe('opening a workbook', () => {
	it('evaluates no formula until the first edit', () => {
		const wb = opened();
		createEditSession(wb);
		expect(value(wb, 1, 1)).toBe(999);
		expect(value(wb, 3, 1)).toBeNull();
	});

	it('trusts stored results on the first edit and computes only what it reaches', () => {
		const wb = opened();
		const session = createEditSession(wb);
		session.setCellValue(0, 0, 0, 5);
		expect(value(wb, 0, 1)).toBe(50);
		// Not reached by the edit: the stored (stale) results stay, as in Excel.
		expect(value(wb, 1, 1)).toBe(999);
		expect(value(wb, 2, 1)).toBe(1000);
		// A formula saved without a result is computed.
		expect(value(wb, 3, 1)).toBe(8);
		session.setCellValue(0, 1, 0, 4);
		expect([value(wb, 1, 1), value(wb, 2, 1), value(wb, 3, 1)]).toEqual([40, 41, 9]);
	});

	it('evaluates a formula typed by the first edit incrementally', () => {
		const wb = opened();
		const session = createEditSession(wb);
		session.setCellInput(0, 4, 0, '=SEQUENCE(2)');
		expect([value(wb, 4, 0), value(wb, 5, 0)]).toEqual([1, 2]);
		expect(value(wb, 1, 1)).toBe(999);
	});

	it('calculates everything on the first edit when stored formulas may spill', () => {
		const wb = opened(true);
		const session = createEditSession(wb);
		session.setCellValue(0, 0, 0, 5);
		expect([value(wb, 0, 1), value(wb, 1, 1), value(wb, 2, 1)]).toEqual([50, 30, 31]);
	});

	it('calculates everything on demand (F9)', () => {
		const wb = opened();
		const session = createEditSession(wb);
		session.calculateNow();
		expect([value(wb, 1, 1), value(wb, 2, 1), value(wb, 3, 1)]).toEqual([30, 31, 5]);
	});
});
