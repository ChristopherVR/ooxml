import { describe, expect, it } from 'vitest';
import { getCell } from '../cells';
import { createWorkbook } from '../workbook';
import { dataTableFormulaText } from './data-table';
import { createEditSession } from './session';

// Inputs A1=3 and A2=4; formulas B4 (=A1*10) and C4 (=A1*A1+A2); input values A5:A7 = 1,2,3.
// The one-variable table covers A4:C7, so its results are B5:C7 and its column input is A1.
function setup() {
	const wb = createWorkbook({ sheets: ['Sheet1'] });
	const s = createEditSession(wb);
	s.setCellValue(0, 0, 0, 3);
	s.setCellValue(0, 1, 0, 4);
	s.setCellInput(0, 3, 1, '=A1*10');
	s.setCellInput(0, 3, 2, '=A1*A1+A2');
	[1, 2, 3].forEach((v, i) => s.setCellValue(0, 4 + i, 0, v));
	s.createDataTable(0, {
		range: { start: { row: 3, col: 0 }, end: { row: 6, col: 2 } },
		colInput: 'A1',
	});
	return { wb, s, sheet: wb.sheets[0]! };
}
const rows = (r: [number, number]) => ({ start: r[0], end: r[1] });
const spanOf = (t: { ref: { start: { row: number }; end: { row: number } } } | undefined) =>
	t && rows([t.ref.start.row, t.ref.end.row]);

describe('data tables on row and column edits', () => {
	it('shifts the table and its input cell for an insert before it', () => {
		const { wb, s, sheet } = setup();
		s.insertRows(0, 0, 2);
		expect(sheet.dataTables).toHaveLength(1);
		expect(spanOf(sheet.dataTables?.[0])).toEqual({ start: 6, end: 8 });
		expect(sheet.dataTables?.[0]?.r1).toBe('A3');
		expect([6, 7, 8].map((r) => getCell(wb.sheets[0]!, r, 1)?.value)).toEqual([10, 20, 30]);
		expect(dataTableFormulaText(sheet, 6, 1)).toBe('{=TABLE(,A3)}');
		s.undo();
		expect(spanOf(sheet.dataTables?.[0])).toEqual({ start: 4, end: 6 });
		expect(sheet.dataTables?.[0]?.r1).toBe('A1');
	});

	it('leaves the table alone for an insert after it', () => {
		const { sheet, s } = setup();
		const before = structuredClone(sheet.dataTables);
		s.insertRows(0, 10, 3);
		expect(sheet.dataTables).toEqual(before);
	});

	it('widens the table for an insert inside its results', () => {
		const { sheet, s, wb } = setup();
		s.insertRows(0, 5, 1);
		expect(spanOf(sheet.dataTables?.[0])).toEqual({ start: 4, end: 7 });
		expect(sheet.dataTables?.[0]?.r1).toBe('A1');
		// The rows around the insertion keep their results.
		expect(getCell(wb.sheets[0]!, 4, 1)?.value).toBe(10);
		expect(getCell(wb.sheets[0]!, 7, 1)?.value).toBe(30);
	});

	it('shifts a table sideways for a column insert and moves a column input', () => {
		const { sheet, s } = setup();
		s.insertColumns(0, 0, 1);
		const table = sheet.dataTables?.[0];
		expect(table?.ref.start.col).toBe(2);
		expect(table?.ref.end.col).toBe(3);
		expect(table?.r1).toBe('B1');
	});

	it('trims the table when some result rows are deleted', () => {
		const { sheet, s } = setup();
		s.deleteRows(0, 5, 1);
		expect(spanOf(sheet.dataTables?.[0])).toEqual({ start: 4, end: 5 });
	});

	it('marks a deleted input cell as #REF! and stops recalculating', () => {
		const { sheet, s, wb } = setup();
		s.deleteRows(0, 0, 1);
		const table = sheet.dataTables?.[0];
		expect(table?.del1).toBe(true);
		expect(spanOf(table)).toEqual({ start: 3, end: 5 });
		expect(dataTableFormulaText(sheet, 3, 1)).toBe('{=TABLE(,#REF!)}');
		// The results stay as the last computed values.
		expect(getCell(wb.sheets[0]!, 3, 1)?.value).toBe(10);
		s.undo();
		expect(sheet.dataTables?.[0]?.del1).toBeUndefined();
		expect(sheet.dataTables?.[0]?.r1).toBe('A1');
	});

	it('removes the table when its header row is deleted but keeps the values', () => {
		const { sheet, s, wb } = setup();
		s.deleteRows(0, 3, 1);
		expect(sheet.dataTables).toBeUndefined();
		expect(getCell(wb.sheets[0]!, 3, 1)?.value).toBe(10);
	});

	it('removes the table when all of its results are deleted', () => {
		const { sheet, s } = setup();
		s.deleteRows(0, 4, 3);
		expect(sheet.dataTables).toBeUndefined();
	});
});
