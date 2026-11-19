import { describe, expect, it } from 'vitest';
import { cellKey } from '../address';
import { putCell } from '../cells';
import { createWorkbook } from '../workbook';
import { accessibleWindow, describeCell } from './accessibility';

const salesSheet = () => {
	const sheet = createWorkbook().sheets[0]!;
	putCell(sheet, 0, 0, { value: 'Quarter' });
	putCell(sheet, 0, 1, { value: 'Sales' });
	putCell(sheet, 1, 0, { value: 'Q1' });
	putCell(sheet, 1, 1, { value: 100 });
	putCell(sheet, 2, 1, { value: 200, formula: 'B2*2' });
	sheet.tables.push({
		id: 1,
		name: 'SalesTable',
		displayName: 'SalesTable',
		range: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
		headerRow: true,
		totalsRow: false,
		columns: [{ name: 'Quarter' }, { name: 'Sales' }],
	});
	return sheet;
};

describe('describeCell', () => {
	it('names the table column of a cell below the header row', () => {
		const sheet = salesSheet();
		expect(describeCell(sheet, { row: 1, col: 1 })).toEqual({ ref: 'B2', header: 'Sales' });
		// The header row itself and cells outside the table have no column context.
		expect(describeCell(sheet, { row: 0, col: 1 })).toEqual({ ref: 'B1' });
		expect(describeCell(sheet, { row: 1, col: 4 })).toEqual({ ref: 'E2' });
	});

	it('gives the formula as the formula bar shows it, array formulas in braces', () => {
		const sheet = salesSheet();
		expect(describeCell(sheet, { row: 2, col: 1 }).formula).toBe('=B2*2');
		putCell(sheet, 5, 5, {
			value: 1,
			formula: 'A1:A2*2',
			arrayRange: { start: { row: 5, col: 5 }, end: { row: 6, col: 5 } },
		});
		expect(describeCell(sheet, { row: 5, col: 5 }).formula).toBe('{=A1:A2*2}');
	});

	it("includes the cell's comment text", () => {
		const sheet = salesSheet();
		sheet.comments.push({ address: { row: 1, col: 0 }, author: 'Ann', text: 'Check this' });
		expect(describeCell(sheet, { row: 1, col: 0 })).toEqual({
			ref: 'A2',
			header: 'Quarter',
			comment: 'Check this',
		});
	});
});

describe('accessibleWindow', () => {
	const visible = { frozenRows: [0], frozenCols: [], rows: [4, 5, 6], cols: [2, 3] };

	it('joins the frozen and scrolling panes and keeps the active cell reachable', () => {
		const window = accessibleWindow(undefined, visible, { row: 40, col: 0 });
		expect(window.rows).toEqual([0, 4, 5, 6, 40]);
		expect(window.cols).toEqual([0, 2, 3]);
	});

	it('spans a merge from its anchor and leaves its covered cells out', () => {
		const sheet = createWorkbook().sheets[0]!;
		sheet.merges.push({ start: { row: 4, col: 2 }, end: { row: 5, col: 3 } });
		sheet.merges.push({ start: { row: 90, col: 2 }, end: { row: 91, col: 3 } });
		const window = accessibleWindow(sheet, visible);
		expect([...window.spans]).toEqual([[cellKey(4, 2), { rows: 2, cols: 2 }]]);
		expect([...window.covered].sort((a, b) => a - b)).toEqual(
			[cellKey(4, 3), cellKey(5, 2), cellKey(5, 3)].sort((a, b) => a - b),
		);
	});
});
