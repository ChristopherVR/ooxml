import { describe, expect, it } from 'vitest';
import { createEditSession } from '../../edit/session';
import { createWorkbook } from '../../workbook';
import { engine, get, set } from '../test-helpers';

const pivotBook = () => {
	const wb = createWorkbook({ sheets: ['Data', 'Summary'] });
	const edit = createEditSession(wb);
	const rows = [
		['Category', 'Region', 'Amount'],
		['Fruit', 'East', 10],
		['Fruit', 'West', 20],
		['Veg', 'East', 5],
		['Veg', 'West', 15],
		['Fruit', 'East', 7],
	];
	rows.forEach((row, r) => row.forEach((value, c) => edit.setCellInput(0, r, c, String(value))));
	edit.createPivotTable(
		1,
		{ row: 0, col: 0 },
		{
			source: 'Data!A1:C6',
			rowField: 'Category',
			columnField: 'Region',
			valueField: 'Amount',
		},
	);
	return wb;
};

const run = (formula: string) => {
	const wb = pivotBook();
	set(wb, 'Summary!H1', `=${formula}`);
	engine(wb).recalculateAll();
	return get(wb, 'Summary!H1');
};

describe('GETPIVOTDATA', () => {
	it('returns the grand total, item totals and intersections', () => {
		expect(run('GETPIVOTDATA("Amount",A3)')).toBe(57);
		expect(run('GETPIVOTDATA("Sum of Amount",A3,"Category","Fruit")')).toBe(37);
		expect(run('GETPIVOTDATA("Amount",A3,"Region","East")')).toBe(22);
		expect(run('GETPIVOTDATA("Amount",A3,"Category","Fruit","Region","East")')).toBe(17);
	});
	it('returns #REF! for unknown fields, items, data fields or non-pivot cells', () => {
		const ref = { error: '#REF!' };
		expect(run('GETPIVOTDATA("Amount",A3,"Category","Meat")')).toEqual(ref);
		expect(run('GETPIVOTDATA("Amount",A3,"Nope","x")')).toEqual(ref);
		expect(run('GETPIVOTDATA("Cost",A3)')).toEqual(ref);
		expect(run('GETPIVOTDATA("Amount",Z50)')).toEqual(ref);
	});
});
