// Timelines over PivotTables with more date records than a call's argument limit (about
// 125,000 in Node): the date range must not pass one argument per record.
import { describe, expect, it } from 'vitest';
import { getCell, putCell } from '../cells';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';

describe('timelines over many date records', () => {
	const ROWS = 200_000;
	it('insert with the range of the earliest and latest dates', () => {
		const workbook = createWorkbook({ sheets: ['Data', 'Summary'] });
		const edit = createEditSession(workbook, { autoRowHeight: false });
		['Category', 'Amount', 'Day'].forEach((value, col) => edit.setCellInput(0, 0, col, value));
		edit.setCellInput(0, 1, 2, '2025-01-15');
		const data = workbook.sheets[0]!;
		const styleId = getCell(data, 1, 2)!.styleId!;
		// Every record falls in 2025 except the earliest and latest, which sit mid-column.
		const day = (row: number) =>
			row === 70_000 ? 43_626 : row === 130_000 ? 47_912 : 45_658 + (row % 365);
		for (let row = 1; row <= ROWS; row++) {
			putCell(data, row, 0, { value: row % 2 ? 'Fruit' : 'Veg' });
			putCell(data, row, 1, { value: 1 });
			putCell(data, row, 2, { value: day(row), styleId });
		}
		edit.createPivotTable(
			1,
			{ row: 0, col: 0 },
			{ source: `Data!A1:C${ROWS + 1}`, rowField: 'Category', valueField: 'Amount' },
		);
		expect(workbook.sheets[1]!.pivotTables![0]!.dateFields).toEqual(['Day']);
		expect(edit.insertTimeline(1, { pivot: 'PivotTable1', field: 'Day' })).toBe('Day');
		// 43626 is 2019-06-10 and 47912 is 2031-03-05: the bounds span their whole years.
		expect(workbook.slicers!.timelineCaches[0]!.bounds).toEqual({
			start: '2019-01-01T00:00:00',
			end: '2032-01-01T00:00:00',
		});
		expect(workbook.slicers!.timelines[0]!.scrollPosition).toBe('2019-01-01T00:00:00');
	});
});
