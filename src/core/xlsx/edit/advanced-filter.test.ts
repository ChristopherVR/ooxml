import { describe, expect, it } from 'vitest';
import { parseRange } from '../address';
import { getCell } from '../cells';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';

function setupSales() {
	const workbook = createWorkbook();
	const session = createEditSession(workbook, { autoRowHeight: false });
	session.setRangeValues(0, { row: 0, col: 0 }, [
		['Region', 'Rep', 'Sales', 'Notes'],
		['East', 'Alice', 100, 'VIP'],
		['East', 'Bob', 250, ''],
		['West', 'Cara', 50, 'New'],
		['West', 'Dan', 300, 'VIP'],
		['East', 'Alice', 100, 'VIP'],
		['North', 'Eve', 180, null],
	]);
	return { workbook, session, sheet: workbook.sheets[0]! };
}

describe('applyAdvancedFilter', () => {
	it('filters in place with AND across columns and OR across rows', () => {
		const { session, sheet } = setupSales();
		// Criteria: (Region = East AND Sales >= 200) OR (Region = West AND Sales < 100)
		session.setRangeValues(0, { row: 0, col: 5 }, [
			['Region', 'Sales'],
			['East', '>=200'],
			['West', '<100'],
		]);

		const res = session.applyAdvancedFilter(0, {
			listRange: parseRange('A1:D7')!,
			criteriaRange: parseRange('F1:G3')!,
			action: 'filter-in-place',
		});

		expect(res.matched).toBe(2);
		expect(res.total).toBe(6);
		expect(sheet.rowInfo.get(0)?.hidden).toBeUndefined();
		expect(sheet.rowInfo.get(1)?.hidden).toBe(true); // East, 100
		expect(sheet.rowInfo.get(2)?.hidden).toBeUndefined(); // East, 250 -> match
		expect(sheet.rowInfo.get(3)?.hidden).toBeUndefined(); // West, 50 -> match
		expect(sheet.rowInfo.get(4)?.hidden).toBe(true); // West, 300
		expect(sheet.rowInfo.get(5)?.hidden).toBe(true); // East, 100
		expect(sheet.rowInfo.get(6)?.hidden).toBe(true); // North, 180
	});

	it('supports comparison operators =, <>, >, >=, <, <= and duplicate criteria headers', () => {
		const { session, sheet } = setupSales();
		// Criteria: Sales > 100 AND Sales <= 250 (matches 250 and 180)
		session.setRangeValues(0, { row: 0, col: 5 }, [
			['Sales', 'Sales'],
			['>100', '<=250'],
		]);

		const res = session.applyAdvancedFilter(0, {
			listRange: parseRange('A1:D7')!,
			criteriaRange: parseRange('F1:G2')!,
			action: 'filter-in-place',
		});

		expect(res.matched).toBe(2);
		expect(sheet.rowInfo.get(2)?.hidden).toBeUndefined(); // 250
		expect(sheet.rowInfo.get(6)?.hidden).toBeUndefined(); // 180
		expect(sheet.rowInfo.get(1)?.hidden).toBe(true); // 100
		expect(sheet.rowInfo.get(4)?.hidden).toBe(true); // 300
	});

	it('supports wildcards (* and ?) and blank/non-blank (= and <>) criteria', () => {
		const { session, sheet } = setupSales();
		// Wildcard: Rep matches "A*" OR Notes is blank "="
		session.setRangeValues(0, { row: 0, col: 5 }, [
			['Rep', 'Notes'],
			['A*', null],
			[null, '='],
		]);

		const res = session.applyAdvancedFilter(0, {
			listRange: parseRange('A1:D7')!,
			criteriaRange: parseRange('F1:G3')!,
			action: 'filter-in-place',
		});

		// Matches: row 1 (Alice), row 2 (blank Notes), row 5 (Alice), row 6 (null Notes)
		expect(res.matched).toBe(4);
		expect(sheet.rowInfo.get(1)?.hidden).toBeUndefined();
		expect(sheet.rowInfo.get(2)?.hidden).toBeUndefined();
		expect(sheet.rowInfo.get(3)?.hidden).toBe(true);
		expect(sheet.rowInfo.get(4)?.hidden).toBe(true);
		expect(sheet.rowInfo.get(5)?.hidden).toBeUndefined();
		expect(sheet.rowInfo.get(6)?.hidden).toBeUndefined();

		// Single-character wildcard "?" and non-blank "<>"
		session.setRangeValues(0, { row: 5, col: 5 }, [
			['Rep', 'Notes'],
			['???', '<>'],
		]);
		const res2 = session.applyAdvancedFilter(0, {
			listRange: parseRange('A1:D7')!,
			criteriaRange: parseRange('F6:G7')!,
			action: 'filter-in-place',
		});
		// 3-letter names with non-blank Notes: Dan (row 4). Bob has blank Notes, Eve has null Notes.
		expect(res2.matched).toBe(1);
		expect(sheet.rowInfo.get(4)?.hidden).toBeUndefined();
		expect(sheet.rowInfo.get(2)?.hidden).toBe(true);
		expect(sheet.rowInfo.get(6)?.hidden).toBe(true);
	});

	it('copies all columns and matching rows to a new location', () => {
		const { session, sheet } = setupSales();
		session.setRangeValues(0, { row: 0, col: 5 }, [['Region'], ['East']]);

		const res = session.applyAdvancedFilter(0, {
			listRange: parseRange('A1:D7')!,
			criteriaRange: parseRange('F1:F2')!,
			action: 'copy-to-location',
			copyTo: { row: 9, col: 0 },
		});

		expect(res.matched).toBe(3);
		expect(res.range).toEqual(parseRange('A10:D13'));
		expect(getCell(sheet, 9, 0)?.value).toBe('Region');
		expect(getCell(sheet, 9, 1)?.value).toBe('Rep');
		expect(getCell(sheet, 10, 1)?.value).toBe('Alice');
		expect(getCell(sheet, 11, 1)?.value).toBe('Bob');
		expect(getCell(sheet, 12, 1)?.value).toBe('Alice');
	});

	it('supports uniqueOnly deduplication for both filter-in-place and copy-to-location', () => {
		const { session, sheet } = setupSales();
		// No criteria + uniqueOnly: drops duplicate row 5 (East, Alice, 100, VIP)
		const inPlace = session.applyAdvancedFilter(0, {
			listRange: parseRange('A1:D7')!,
			action: 'filter-in-place',
			uniqueOnly: true,
		});
		expect(inPlace.matched).toBe(5);
		expect(sheet.rowInfo.get(1)?.hidden).toBeUndefined();
		expect(sheet.rowInfo.get(5)?.hidden).toBe(true);

		const copied = session.applyAdvancedFilter(0, {
			listRange: parseRange('A1:D7')!,
			action: 'copy-to-location',
			copyTo: { row: 9, col: 0 },
			uniqueOnly: true,
		});
		expect(copied.matched).toBe(5);
		expect(copied.range).toEqual(parseRange('A10:D15'));
	});

	it('extracts a requested subset of headers when copyTo has pre-populated headers', () => {
		const { session, sheet } = setupSales();
		session.setRangeValues(0, { row: 0, col: 5 }, [['Sales'], ['>=180']]);
		// Pre-populate desired output headers in reverse order at A10:B10
		session.setRangeValues(0, { row: 9, col: 0 }, [['Sales', 'Rep']]);

		const res = session.applyAdvancedFilter(0, {
			listRange: parseRange('A1:D7')!,
			criteriaRange: parseRange('F1:F2')!,
			action: 'copy-to-location',
			copyTo: parseRange('A10:B10')!,
		});

		expect(res.matched).toBe(3);
		expect(res.range).toEqual(parseRange('A10:B13'));
		expect(getCell(sheet, 9, 0)?.value).toBe('Sales');
		expect(getCell(sheet, 9, 1)?.value).toBe('Rep');
		expect(getCell(sheet, 10, 0)?.value).toBe(250);
		expect(getCell(sheet, 10, 1)?.value).toBe('Bob');
		expect(getCell(sheet, 11, 0)?.value).toBe(300);
		expect(getCell(sheet, 11, 1)?.value).toBe('Dan');
		expect(getCell(sheet, 12, 0)?.value).toBe(180);
		expect(getCell(sheet, 12, 1)?.value).toBe('Eve');
	});

	it('supports full undo and redo for both actions', () => {
		const { session, sheet } = setupSales();
		session.setRangeValues(0, { row: 0, col: 5 }, [['Region'], ['West']]);

		// Filter in place + undo/redo
		session.applyAdvancedFilter(0, {
			listRange: parseRange('A1:D7')!,
			criteriaRange: parseRange('F1:F2')!,
			action: 'filter-in-place',
		});
		expect(sheet.rowInfo.get(1)?.hidden).toBe(true);
		expect(sheet.rowInfo.get(3)?.hidden).toBeUndefined();

		expect(session.undo()).toBe(true);
		expect(sheet.rowInfo.get(1)?.hidden).toBeUndefined();

		expect(session.redo()).toBe(true);
		expect(sheet.rowInfo.get(1)?.hidden).toBe(true);

		// Copy to location + undo/redo
		session.applyAdvancedFilter(0, {
			listRange: parseRange('A1:D7')!,
			criteriaRange: parseRange('F1:F2')!,
			action: 'copy-to-location',
			copyTo: { row: 9, col: 0 },
		});
		expect(getCell(sheet, 10, 1)?.value).toBe('Cara');
		expect(getCell(sheet, 11, 1)?.value).toBe('Dan');

		expect(session.undo()).toBe(true);
		expect(getCell(sheet, 9, 0)).toBeUndefined();
		expect(getCell(sheet, 10, 1)).toBeUndefined();

		expect(session.redo()).toBe(true);
		expect(getCell(sheet, 10, 1)?.value).toBe('Cara');
		expect(getCell(sheet, 11, 1)?.value).toBe('Dan');
	});

	it('throws when copy-to-location omits copyTo', () => {
		const { session } = setupSales();
		expect(() =>
			session.applyAdvancedFilter(0, {
				listRange: parseRange('A1:D7')!,
				action: 'copy-to-location',
			}),
		).toThrow(RangeError);
	});
});
